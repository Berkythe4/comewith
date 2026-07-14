// Tests for assets/import-engine.js — runs the real Come With 7-11 exports
// through the pipeline and asserts the numbers that were hand-verified on prod
// during the 2026-07-13 import. Run: node scripts/test_import_engine.mjs
import { readFileSync } from 'node:fs';
import { parseCsv, classifyExport, buildImportPlan, matchName, stripHash } from '../assets/import-engine.js';

const DIR = new URL('../events/come-with/7-11/', import.meta.url);
const load = (f) => parseCsv(readFileSync(new URL(f, DIR), 'utf8'));

let pass = 0, fail = 0;
const eq = (label, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log(`${ok ? '  ok' : 'FAIL'}  ${label}${ok ? '' : `  → got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`}`);
};

// ---- parseCsv ----
{
  const { headers, rows } = parseCsv('a,b\n"x, ""q""",2\r\n,3\n');
  eq('parseCsv quoted fields', rows[0], { a: 'x, "q"', b: '2' });
  eq('parseCsv headers', headers, ['a', 'b']);
  eq('parseCsv keeps row with only second cell', rows[1], { a: '', b: '3' });
}

// ---- classification of the real exports ----
const full = load('20260711-ComeWith-list (3).csv');
const basic = load('Come_With_7-11.csv');
const scans = load('20260711-ComeWith-scandata.csv');
const gl = load('RA_guestlist_Come_With_7-11.csv');
const pf = load('ComeWith_7-13_guests_partiful.csv');
eq('classify full RA list', classifyExport(full.headers), 'ra_tickets');
eq('classify basic RA list (no email col)', classifyExport(basic.headers), 'ra_tickets');
eq('classify scan data', classifyExport(scans.headers), 'ra_scans');
eq('classify RA guestlist', classifyExport(gl.headers), 'ra_guestlist');
eq('classify Partiful', classifyExport(pf.headers), 'partiful');
eq('classify unknown', classifyExport(['Foo', 'Bar']), null);

// ---- name matching ----
eq('stripHash', stripHash('Lunaera #2'), 'Lunaera');
eq('match exact ci', matchName('emma stroble', [{ name: 'Emma Stroble' }])?.via, 'exact');
eq('match single token → first name', matchName('Lila', [{ name: 'Lila Bey' }])?.via, 'first name');
eq('match stage name via email handle', matchName('Knostalgia', [{ name: 'Knyckolas Sutherland', email: 'knostalgiamusic@gmail.com' }])?.via, 'email handle');
eq('no match on +1 pseudo-name', matchName("Knostalgia's +1", [{ name: 'Knyckolas Sutherland', email: 'knostalgiamusic@gmail.com' }]), null);
eq('no match Martin vs Just Martin', matchName('Martin', [{ name: 'Just Martin' }]), null);

const event = { id: 'ev-711', slug: 'come-with-7-11', series: 'Come With Parties', name: 'Come With 7-11' };
const HOSTS = ['Keith Berkman'];
const files = [
  { name: 'list(3).csv', type: 'ra_tickets', rows: full.rows },
  { name: 'basic.csv', type: 'ra_tickets', rows: basic.rows },     // duplicate export, fewer columns
  { name: 'scandata.csv', type: 'ra_scans', rows: scans.rows },
  { name: 'guestlist.csv', type: 'ra_guestlist', rows: gl.rows },
  { name: 'partiful.csv', type: 'partiful', rows: pf.rows },
];

// ---- full drop on an empty database (what a from-scratch import would do) ----
{
  const plan = buildImportPlan({ files, event, hostNames: HOSTS,
    db: { guests: [], eventGeaGuestIds: [], eventTicketKeys: [], subscribers: [], segments: [] } });
  const ra = plan.ticketing.filter(t => t.source === 'resident_advisor');
  const pfx = plan.ticketing.filter(t => t.source === 'partiful');
  eq('RA ticket rows (merged across both exports, keyed on barcode)', ra.length, 37);
  eq('RA rows carry emails from the richer export', ra.every(t => t.guestKey || t.guestId), true);
  eq('Partiful-only tickets after dedupe + host exclusion', pfx.length, 13);
  eq('attendance = unique scanned barcodes', plan.eventUpdate?.total_attendance, 27);
  eq('event marked completed', plan.eventUpdate?.status, 'completed');
  eq('attendance links: 24 buyers + 15 guestlist + 13 partiful + 4 door', plan.gea.length, 56);
  eq('door walk-ins created', plan.gea.filter(g => g.source === 'ra_door').map(g => g.name).sort(),
     ['Erika Scott', 'Garth', 'Kyle', 'Martin']);
  eq('"Lunaera #2" collapses into Lunaera (no bogus guest)',
     plan.guestsToCreate.some(g => /#\d/.test(g.full_name)), false);
  eq('new subscribers: 24 buyer emails + 11 guestlist emails', plan.subscribers.length, 35);
  eq('brand segment for a parties event', plan.brandSeg, 'come_with');
  eq('event segment = slug', plan.eventSeg, 'come-with-7-11');
  eq('host excluded from tickets', plan.report.skips.some(s => s.includes('Keith Berkman')), true);
  const dd = plan.report.dedupes.map(d => d.name.split(' (')[0]).sort();
  eq('Partiful dedupes vs RA', dd, ['Kyle', 'Lila', 'Liz McQuillan', 'Marc', 'Steve', 'Victoriarose Vargas', 'emma stroble'].concat(['Knostalgia']).sort());
  eq('attended flags: 8 RA ticket barcodes scanned', ra.filter(t => t.attended === true).length, 8);
  // Person-level scanned-in: 7 buyers + Victoriarose (guestlist barcode collapses onto her
  // name) + 9 scanned guest-list names + 4 door walk-ins = 21 people through the door.
  eq('scanned-in people', plan.gea.filter(g => g.attended === true).length, 21);
  eq('listed but never scanned', plan.gea.filter(g => g.attended === false).length, 21);
  eq('unknown (Partiful, no check-in data)', plan.gea.filter(g => g.attended == null).length, 14);
  const gAtt = (n) => plan.gea.find(g => g.name === n)?.attended;
  eq('Liz McQuillan scanned', gAtt('Liz McQuillan'), true);
  eq('Rohon Nandi never scanned', gAtt('Rohon Nandi'), false);
  eq('KRNeY (guest list) scanned', gAtt('KRNeY'), true);
  eq('Marlo (Partiful) unknown', gAtt('Marlo'), undefined);
  eq('geaAttended update list covers everyone determinable', plan.geaAttended.length, 42);
  // Admissions reconcile with the door count: parties roll up (Lunaera ×6, Liz ×2).
  eq('scanned admissions sum = 27 (the door count)', plan.geaAttended.reduce((s, a) => s + (a.scans || 0), 0), 27);
  eq('Lunaera party admissions', plan.gea.find(g => g.name === 'Lunaera')?.scan_count, 6);
  eq('Liz two barcodes scanned', plan.gea.find(g => g.name === 'Liz McQuillan')?.scan_count, 2);
}

// ---- prod-like state: consent + existing-customer linking ----
{
  const db = {
    guests: [
      { id: 'g-chad', full_name: 'Chad Hernandez', email: 'chaddercheesy@gmail.com', opted_in_mailing: false },
      { id: 'g-moody', full_name: 'Alexander Moody', email: 'alex.imoody@gmail.com', opted_in_mailing: true },
      { id: 'g-liz', full_name: 'Liz McQuillan', email: 'emcquillan@gmail.com', opted_in_mailing: true },
    ],
    eventGeaGuestIds: [],
    eventTicketKeys: [],
    subscribers: [
      { id: 's-chad', email: 'chaddercheesy@gmail.com', status: 'unsubscribed' },
      { id: 's-moody', email: 'alex.imoody@gmail.com', status: 'subscribed' },
      { id: 's-liz', email: 'emcquillan@gmail.com', status: 'subscribed' },
    ],
    segments: [],
  };
  const plan = buildImportPlan({ files, event, hostNames: HOSTS, db });
  eq('unsubscribed email NOT re-subscribed', plan.subscribers.some(s => s.email === 'chaddercheesy@gmail.com'), false);
  eq('…and flagged for the report', plan.report.flags.some(f => f.includes('chaddercheesy') && f.includes('NOT re-subscribed')), true);
  eq('existing guests not duplicated', plan.guestsToCreate.some(g => ['chaddercheesy@gmail.com', 'emcquillan@gmail.com'].includes(g.email)), false);
  eq('Partiful Alexander Moody links to his existing record', plan.guestsToCreate.some(g => g.full_name === 'Alexander Moody'), false);
  const moodyGea = plan.gea.find(g => g.name === 'Alexander Moody');
  eq('…with an attendance link on the existing guest id', moodyGea?.guestId, 'g-moody');
  eq('…and his email queued for event+brand segments', plan.segmentEmails.has('alex.imoody@gmail.com'), true);
}

// ---- re-run (everything already imported) is a no-op ----
{
  const firstRun = buildImportPlan({ files, event, hostNames: HOSTS,
    db: { guests: [], eventGeaGuestIds: [], eventTicketKeys: [], subscribers: [], segments: [] } });
  const guests = firstRun.guestsToCreate.map((g, i) => ({ id: 'g' + i, full_name: g.full_name, email: g.email, opted_in_mailing: g.opted_in_mailing }));
  const db = {
    guests,
    eventGeaGuestIds: guests.map(g => g.id),
    eventTicketKeys: firstRun.ticketing.map(t => t.source + '|' + t.external_id),
    subscribers: firstRun.subscribers.map((s, i) => ({ id: 's' + i, email: s.email, status: 'subscribed' })),
    segments: [],
  };
  const rerun = buildImportPlan({ files, event, hostNames: HOSTS, db });
  eq('re-run: no new guests', rerun.guestsToCreate.length, 0);
  eq('re-run: no new tickets', rerun.ticketing.length, 0);
  eq('re-run: no new attendance links', rerun.gea.length, 0);
  eq('re-run: no new subscribers', rerun.subscribers.length, 0);
  eq('re-run: attendance recomputes identically', rerun.eventUpdate?.total_attendance, 27);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
