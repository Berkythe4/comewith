// Import engine — pure logic for the event hub "Import event exports" flow.
// No DOM, no Supabase: dashboard.html parses the dropped CSVs with parseCsv(),
// classifies them with classifyExport(), fetches current DB state, and hands
// everything to buildImportPlan(). The plan lists exactly what to insert/update
// plus a human report (dedupes, flags, skips); dashboard.html executes it.
// Keeping this pure lets scripts/test_import_engine.mjs run the whole pipeline
// against the real Come With 7-11 exports and assert the known-good numbers.
//
// Standing rules encoded here (from the 2026-07-13 Come With 7-11 import):
//  - dedupe guests/subscribers by lower(email); ticketing is keyed on RA BARCODE
//    (order numbers repeat across multi-ticket orders)
//  - attendance = unique barcodes with ScanCount > 0 (RA scan-data export)
//  - Partiful "Going" = ticket sold AND customer record (no emails in Partiful
//    exports) — deduped by NAME against RA tickets/guestlist/scans/existing
//    guests, host excluded; "Maybe" never counts
//  - door walk-ins (scanned barcodes not on any ticket list) become customers;
//    "Name #2"-style RA quantity-overflow rows collapse into the base name
//  - NEVER (re-)subscribe an email that is unsubscribed or whose guest record
//    has opted_in_mailing = false; everyone else with an email is subscribed
//  - every subscribed attendee gets BOTH the per-event segment (event slug) and
//    the brand segment (dance_infusion for Dance Infusion, else come_with)

// ---- CSV ----

export function parseCsv(text) {
  text = String(text || '').replace(/^﻿/, '');
  const rows = [];
  let field = '', row = [], inQ = false;
  const pushF = () => { row.push(field); field = ''; };
  const pushR = () => { pushF(); if (row.some(c => c.trim() !== '')) rows.push(row); row = []; };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ',') pushF();
    else if (ch === '\n') pushR();
    else if (ch === '\r') { if (text[i + 1] === '\n') i++; pushR(); }
    else field += ch;
  }
  if (field !== '' || row.length) pushR();
  if (!rows.length) return { headers: [], rows: [] };
  const headers = rows[0].map(h => h.trim());
  const out = rows.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => { o[h] = (r[i] ?? '').trim(); });
    return o;
  });
  return { headers, rows: out };
}

// ---- Classification (by header signature) ----

export function classifyExport(headers) {
  const h = (headers || []).map(x => String(x).trim().toLowerCase());
  const has = (n) => h.includes(n);
  if (has('barcode') && has('scancount')) return 'ra_scans';
  if (has('barcode') && has('billing name')) return 'ra_tickets';
  if (has('name') && has('status') && has('rsvp date')) return 'partiful';
  if (has('name') && has('email') && has('quantity')) return 'ra_guestlist';
  return null;
}

export const TYPE_LABELS = {
  ra_tickets: 'RA ticket list', ra_scans: 'RA scan data',
  ra_guestlist: 'RA guest list', partiful: 'Partiful RSVPs',
};

// ---- Names ----

export function normName(s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); }
// RA appends " #2", " #3"… when a guest-list entry has quantity > 1 — same person.
export function stripHash(s) { return String(s || '').replace(/\s*#\d+\s*$/, '').trim(); }
const squash = (s) => normName(s).replace(/[^a-z0-9]/g, '');

// Match a name against candidates: exact; single-token vs first name; first name
// + last-name initial; or the squashed name inside a candidate's email handle
// (catches stage names like Knostalgia ↔ knostalgiamusic@…). Returns
// { name, via } for the report, or null. Every hit is surfaced for human review.
export function matchName(name, candidates) {
  const n = normName(stripHash(name));
  if (!n) return null;
  const nt = n.split(' ');
  for (const c of candidates) {
    const cn = normName(stripHash(c.name));
    if (!cn) continue;
    if (cn === n) return { name: c.name, via: 'exact' };
    const ct = cn.split(' ');
    if (nt.length === 1 && ct[0] === n) return { name: c.name, via: 'first name' };
    if (ct.length === 1 && nt[0] === cn) return { name: c.name, via: 'first name' };
    if (nt.length > 1 && ct.length > 1 && nt[0] === ct[0] && nt[nt.length - 1][0] === ct[ct.length - 1][0])
      return { name: c.name, via: 'first name + last initial' };
  }
  const sq = squash(name);
  if (sq.length >= 5) {
    for (const c of candidates) {
      const handle = (c.email || '').split('@')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
      if (handle && handle.includes(sq)) return { name: c.name || c.email, via: 'email handle' };
    }
  }
  return null;
}

export function slugify(s) { return normName(s).replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, ''); }

// ---- Plan builder ----
// files: [{ name, type, rows }]  (already parsed + classified)
// event: { id, slug, series, name }
// db:    { guests: [{id, full_name, email, opted_in_mailing}],
//          eventGeaGuestIds: [uuid], eventTicketKeys: ['source|external_id'],
//          subscribers: [{id, email, status}], segments: [{subscriber_id, segment}] }
// hostNames: names never counted as customers/tickets (event owner etc.)
export function buildImportPlan({ files, event, db, hostNames = [] }) {
  const tag = 'event-import:' + (event.slug || event.id);
  const eventSeg = event.slug || String(event.id);
  const brandSeg = event.series === 'Dance Infusion' ? 'dance_infusion' : 'come_with';

  const guestsByEmail = new Map(), guestsByName = new Map();
  for (const g of (db.guests || [])) {
    if (g.email) guestsByEmail.set(g.email.toLowerCase(), g);
    const k = normName(g.full_name); if (k && !guestsByName.has(k)) guestsByName.set(k, g);
  }
  const onEvent = new Set(db.eventGeaGuestIds || []);
  const existingTicketKeys = new Set(db.eventTicketKeys || []);
  const subsByEmail = new Map();
  for (const s of (db.subscribers || [])) subsByEmail.set((s.email || '').toLowerCase(), s);

  const plan = {
    tag, eventSeg, brandSeg,
    guestsToCreate: [],           // {key, full_name, email, opted_in_mailing, source, notes}
    ticketing: [],                // {external_id, source, ticket_type, amount_paid, quantity, purchased_at, attended, notes, guestId?, guestKey?}
    gea: [],                      // {guestId?|guestKey?, name, amount_spent, ticket_type, quantity, source, purchased_at?}
    attendedTrue: [], attendedFalse: [],   // RA barcodes to flip on existing ticketing rows
    geaAttended: [],              // {guestId?|guestKey?, attended} — person-level scanned-in flags
    eventUpdate: null,            // {total_attendance, status}
    subscribers: [],              // {email, full_name, guestId?|guestKey?}
    segmentEmails: new Set(),     // everyone tied to the event with an email
    report: { detected: [], planned: [], dedupes: [], flags: [], skips: [], unknownFiles: [] },
  };
  const rep = plan.report;

  const newGuestByKey = new Map();
  const newGuestByName = new Map();          // norm(name) -> key of first new guest with that name
  const plannedGeaKeys = new Set();          // guestId/guestKey already getting a gea row
  const geaNameIndex = [];                   // {name, email} of everyone on/joining the event (for dedupe)
  for (const g of (db.guests || [])) if (onEvent.has(g.id)) geaNameIndex.push({ name: g.full_name, email: g.email });

  const ensureGuest = ({ name, email, optIn, notes }) => {
    email = (email || '').trim().toLowerCase() || null;
    if (email && guestsByEmail.has(email)) return { guestId: guestsByEmail.get(email).id, existing: guestsByEmail.get(email) };
    const nk = 'n:' + normName(name), ek = email && 'e:' + email;
    if (ek && newGuestByKey.has(ek)) return { guestKey: ek };
    if (!email) {
      const byName = guestsByName.get(normName(name));
      if (byName) return { guestId: byName.id, existing: byName };
      if (newGuestByKey.has(nk)) return { guestKey: nk };
    }
    const key = ek || nk;
    const g = { key, full_name: String(name || '').trim() || email, email,
                opted_in_mailing: !!(optIn && email), source: tag, notes: notes || null };
    newGuestByKey.set(key, g); plan.guestsToCreate.push(g);
    const nn = normName(name); if (nn && !newGuestByName.has(nn)) newGuestByName.set(nn, key);
    return { guestKey: key };
  };
  const addGea = (ref, row) => {
    const k = ref.guestId || ref.guestKey;
    // Index the name either way — later dedupe passes (door scans, Partiful)
    // must see this person even when their attendance row already exists.
    geaNameIndex.push({ name: row.name, email: row.email || ref.existing?.email });
    if (ref.guestId && onEvent.has(ref.guestId)) return false;
    if (plannedGeaKeys.has(k)) return false;
    plannedGeaKeys.add(k);
    plan.gea.push({ ...row, guestId: ref.guestId, guestKey: ref.guestKey });
    return true;
  };

  // -- merge files by type --
  const ticketRows = new Map();   // barcode -> row (prefer the row that has an email)
  const scanRows = new Map();     // barcode -> {name, count, time, ticketType}
  const glRows = [], pfRows = [];
  for (const f of files) {
    if (!f.type) { rep.unknownFiles.push(f.name); continue; }
    rep.detected.push({ file: f.name, type: f.type, rows: f.rows.length });
    for (const r of f.rows) {
      if (f.type === 'ra_tickets') {
        const bc = r['Barcode']; if (!bc) continue;
        const prev = ticketRows.get(bc);
        if (!prev || (!prev['Email'] && r['Email'])) ticketRows.set(bc, r);
      } else if (f.type === 'ra_scans') {
        const bc = r['Barcode']; if (!bc) continue;
        const cur = scanRows.get(bc);
        const count = Number(r['ScanCount'] || 0);
        if (!cur || count > cur.count) scanRows.set(bc, { name: r['Name'] || '', count, time: r['ScanDateTime'] || '', ticketType: r['TicketType'] || '' });
      } else if (f.type === 'ra_guestlist') glRows.push(r);
      else if (f.type === 'partiful') pfRows.push(r);
    }
  }

  // -- 1) RA tickets: ticketing per barcode + guest/gea per unique buyer --
  const buyers = new Map();       // email or name-key -> {name, email, rows:[]}
  let tixNew = 0, tixSkip = 0;
  for (const [bc, r] of ticketRows) {
    const email = (r['Email'] || '').toLowerCase() || null;
    const name = r['Billing name'] || email || '(unknown)';
    const bk = email || 'name:' + normName(name);
    if (!buyers.has(bk)) buyers.set(bk, { name, email, rows: [] });
    buyers.get(bk).rows.push(r);
    if (existingTicketKeys.has('resident_advisor|' + bc)) { tixSkip++; continue; }
    const scan = scanRows.get(bc);
    plan.ticketing.push({
      external_id: bc, source: 'resident_advisor',
      ticket_type: r['Ticket type'] || 'General admission',
      amount_paid: Number(r['Price'] || 0) || 0, quantity: Number(r['Quantity'] || 1) || 1,
      purchased_at: r['Date purchased'] || null,
      attended: scanRows.size ? !!(scan && scan.count > 0) : null,
      notes: 'order ' + (r['Order number'] || '—'), buyerKey: bk,
    });
    tixNew++;
  }
  for (const [, b] of buyers) {
    const ref = ensureGuest({ name: b.name, email: b.email, optIn: true });
    for (const t of plan.ticketing) if (t.buyerKey === ('' + (b.email || 'name:' + normName(b.name)))) { t.guestId = ref.guestId; t.guestKey = ref.guestKey; }
    addGea(ref, {
      name: b.name, email: b.email,
      amount_spent: b.rows.reduce((s, r) => s + (Number(r['Price'] || 0) || 0), 0),
      ticket_type: b.rows[0]['Ticket type'] || 'General admission',
      quantity: b.rows.reduce((s, r) => s + (Number(r['Quantity'] || 1) || 1), 0),
      source: 'resident_advisor',
      purchased_at: b.rows.map(r => r['Date purchased']).filter(Boolean).sort()[0] || null,
    });
    if (b.email) plan.segmentEmails.add(b.email);
  }
  if (ticketRows.size) rep.planned.push(`RA tickets: ${tixNew} new ticket rows (${tixSkip} already imported) across ${buyers.size} buyers`);

  // -- 2) RA guest list --
  let glNew = 0;
  for (const r of glRows) {
    const name = r['Name'], email = (r['Email'] || '').toLowerCase() || null;
    if (!name && !email) continue;
    const ref = ensureGuest({ name, email, optIn: true });
    if (addGea(ref, { name, email, amount_spent: 0, ticket_type: 'comp',
                      quantity: Number(r['Quantity'] || 1) || 1, source: 'guestlist' })) glNew++;
    if (email) plan.segmentEmails.add(email);
  }
  if (glRows.length) rep.planned.push(`Guest list: ${glNew} added to the event (${glRows.length - glNew} already on it)`);

  // -- 3) Scan data: attendance + attended flags + door walk-ins --
  if (scanRows.size) {
    let scanned = 0;
    for (const [bc, s] of scanRows) {
      if (s.count > 0) { scanned++; plan.attendedTrue.push(bc); } else plan.attendedFalse.push(bc);
    }
    plan.eventUpdate = { total_attendance: scanned, status: 'completed' };
    rep.planned.push(`Attendance: ${scanned} scanned in (of ${scanRows.size} barcodes) → total_attendance = ${scanned}, event marked completed`);
    for (const [bc, s] of scanRows) {
      if (s.count <= 0 || ticketRows.has(bc) || existingTicketKeys.has('resident_advisor|' + bc)) continue;
      const base = stripHash(s.name);
      const hit = matchName(base, geaNameIndex);
      if (hit) { if (normName(hit.name) !== normName(base)) rep.dedupes.push({ name: s.name + ' (door scan)', matchedTo: hit.name, via: hit.via }); continue; }
      const existing = guestsByName.get(normName(base));
      const ref = ensureGuest({ name: base, email: null, optIn: false,
        notes: `Door walk-in ${eventSeg}, RA barcode ${bc} scanned ${s.time || '—'}, no email; match later` });
      addGea(ref, { name: base, amount_spent: 0, ticket_type: s.ticketType || 'Guest', quantity: 1, source: 'ra_door' });
      rep.flags.push(existing
        ? `Door walk-in "${base}" linked to existing customer record (no email on file)`
        : `Door walk-in "${base}" created as a customer with NO email (barcode ${bc}) — match later`);
    }

    // Person-level scanned-in flags: resolve every scan barcode to a guest ref
    // (ticket buyers by email, everyone else by name) and OR the scans together.
    const resolveScanRef = (bc, s) => {
      const trow = ticketRows.get(bc);
      let name = stripHash(s.name);
      if (trow) {
        const em = (trow['Email'] || '').toLowerCase();
        if (em) return guestsByEmail.has(em) ? guestsByEmail.get(em).id : (newGuestByKey.has('e:' + em) ? 'e:' + em : null);
        name = trow['Billing name'] || name;
      }
      const nn = normName(name);
      if (guestsByName.has(nn)) return guestsByName.get(nn).id;
      return newGuestByName.get(nn) || null;
    };
    const att = new Map();
    for (const [bc, s] of scanRows) {
      const k = resolveScanRef(bc, s);
      if (k) att.set(k, (att.get(k) || false) || s.count > 0);
    }
    for (const [k, val] of att)
      plan.geaAttended.push(k.startsWith('e:') || k.startsWith('n:') ? { guestKey: k, attended: val } : { guestId: k, attended: val });
    for (const g of plan.gea) {
      const k = g.guestId || g.guestKey;
      if (att.has(k)) g.attended = att.get(k);
    }
  }

  // -- 4) Partiful: Going = ticket sold + customer, deduped by name --
  const going = pfRows.filter(r => (r['Status'] || '').toLowerCase() === 'going');
  const maybes = pfRows.filter(r => (r['Status'] || '').toLowerCase() === 'maybe');
  if (pfRows.length) {
    const hosts = hostNames.filter(Boolean).map(h => ({ name: h }));
    const scanIdx = [...scanRows.values()].map(s => ({ name: stripHash(s.name) }));
    let pfNew = 0, pfSkip = 0;
    for (const r of going) {
      const name = r['Name']; if (!name) continue;
      if (matchName(name, hosts)) { rep.skips.push(`"${name}" is the host — not counted as a ticket`); continue; }
      const hit = matchName(name, geaNameIndex) || matchName(name, scanIdx);
      if (hit) { rep.dedupes.push({ name: name + ' (Partiful)', matchedTo: hit.name, via: hit.via }); continue; }
      const ext = 'partiful:' + slugify(name);
      const plus = r['Is Plus One Of'] || '';
      const existing = guestsByName.get(normName(name));
      const ref = ensureGuest({ name, email: null, optIn: false,
        notes: `Partiful Going RSVP ${eventSeg}, no email in export` + (plus ? `; +1 of ${plus}` : '') });
      if (existingTicketKeys.has('partiful|' + ext)) pfSkip++;
      else {
        plan.ticketing.push({ external_id: ext, source: 'partiful', ticket_type: 'Partiful RSVP',
          amount_paid: 0, quantity: 1, purchased_at: r['RSVP date'] || null, attended: null,
          notes: 'Partiful RSVP Going: ' + name + (plus ? ` (+1 of ${plus})` : ''),
          guestId: ref.guestId, guestKey: ref.guestKey });
        pfNew++;
      }
      addGea(ref, { name, amount_spent: 0, ticket_type: 'Partiful RSVP', quantity: 1, source: 'partiful' });
      if (ref.existing?.email) {
        plan.segmentEmails.add(ref.existing.email.toLowerCase());
        rep.flags.push(`Partiful "${name}" matched existing customer ${ref.existing.full_name} (${ref.existing.email})`);
      }
    }
    rep.planned.push(`Partiful: ${going.length} Going → ${pfNew} new tickets (${rep.dedupes.filter(d => d.name.includes('(Partiful)')).length} already on RA side${pfSkip ? `, ${pfSkip} already imported` : ''}); ${maybes.length} Maybe not counted`);
  }

  // -- 5) Mailing list: subscribe new emails only, honoring every prior opt-out --
  const candidates = new Map();
  for (const [, b] of buyers) if (b.email) candidates.set(b.email, b.name);
  for (const r of glRows) { const e = (r['Email'] || '').toLowerCase(); if (e && !candidates.has(e)) candidates.set(e, r['Name']); }
  let subNew = 0, subHave = 0;
  for (const [email, name] of candidates) {
    const ex = subsByEmail.get(email);
    if (ex) {
      if (ex.status !== 'subscribed') rep.flags.push(`${email} stays ${ex.status} — prior opt-out honored, NOT re-subscribed`);
      else subHave++;
      continue;
    }
    const g = guestsByEmail.get(email);
    if (g && g.opted_in_mailing === false) { rep.flags.push(`${email} (${g.full_name}) has mailing opt-in OFF on their customer record — not subscribed`); continue; }
    plan.subscribers.push({ email, full_name: name, guestId: g?.id,
      guestKey: newGuestByKey.has('e:' + email) ? 'e:' + email : undefined });
    subNew++;
  }
  if (candidates.size) rep.planned.push(`Mailing list: ${subNew} new subscribers (${subHave} already subscribed) → segments "${eventSeg}" + "${brandSeg}"`);

  return plan;
}
