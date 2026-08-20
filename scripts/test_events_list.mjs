// The events list gained four in-place dropdowns (type / stage / status /
// public). A <select> that offers a value the database will reject is worse than
// a text box, because it looks safe — so this asserts the option lists match the
// CHECK constraints on public.events exactly.
//
// The expected values are transcribed from prod on 2026-08-20:
//   events_type_check    party, dance_infusion, production, showcase, gig, growth
//   events_stage_check   idea, planning, confirmed, live, wrapped, reported
//   events_status_check  planning, announced, on_sale, sold_out, completed, cancelled
// If a migration widens one of these, update BOTH the constraint and this list.
//
//   node scripts/test_events_list.mjs        (from the repo root)
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);

const EXPECTED = {
  EV_TYPE_OPTS: ['party', 'dance_infusion', 'production', 'showcase', 'gig', 'growth'],
  EV_STAGE_OPTS: ['idea', 'planning', 'confirmed', 'live', 'wrapped', 'reported'],
  EV_STATUS_OPTS: ['planning', 'announced', 'on_sale', 'sold_out', 'completed', 'cancelled'],
};

for (const [name, want] of Object.entries(EXPECTED)) {
  const m = mod.match(new RegExp('const ' + name + '\\s*=\\s*\\[([\\s\\S]*?)\\];'));
  if (!m) { fail(name + ' is not defined'); continue; }
  const got = [...m[1].matchAll(/\['([^']+)',/g)].map(x => x[1]);
  const missing = want.filter(v => !got.includes(v));
  const extra = got.filter(v => !want.includes(v));
  if (missing.length) fail(name + ' is missing ' + missing.join(', '));
  else if (extra.length) fail(name + ' offers ' + extra.join(', ') + ' which the CHECK constraint rejects');
  else pass(name + ' matches its CHECK constraint (' + got.length + ' values)');
}

// The four columns exist and are editable, not badges.
const table = mod.slice(mod.indexOf('function renderEventsTable()'), mod.indexOf('async function setEventField'));
for (const f of ['type', 'stage', 'status', 'is_public']) {
  if (!table.includes("evSel(e, '" + f + "'")) fail('the ' + f + ' column is not an editable select');
}
if (!/<th>Date<\/th><th>Event<\/th><th>Type<\/th><th>Stage<\/th><th>Status<\/th><th>Public<\/th>/.test(table)) {
  fail('the header does not read Date / Event / Type / Stage / Status / Public');
} else pass('Type and Stage are their own columns, all four are editable selects');

// Twelve columns now, so the empty state has to span twelve.
const cols = (table.match(/<th[ >]/g) || []).length;
const span = (table.match(/colspan="(\d+)"/) || [])[1];
if (String(cols) !== span) fail('the empty-row colspan is ' + span + ' but the table has ' + cols + ' columns');
else pass('empty-state colspan matches the ' + cols + ' columns');

// Growth & Networking went missing from the snapshot because that list was
// hand-written. It must be derived from the data now.
if (/const seriesStr = \[/.test(mod)) fail('the snapshot series breakdown is still a hardcoded array');
else if (!mod.includes('const seriesStr = evSeriesList()')) fail('the snapshot does not use evSeriesList()');
else pass('the series breakdown is derived from the data, so a new series cannot go missing');

if (!mod.includes('const seriesOpts = evSeriesList()')) fail('the series filter is not derived from the data');
else pass('the series filter is derived from the data too');

if (!/EV_SERIES_KNOWN[\s\S]{0,220}Growth & Networking/.test(mod)) fail('Growth & Networking is not in the known-series list');
else pass('Growth & Networking is a first-class series');

// The write path must coerce is_public to a real boolean, not the string "true".
const setter = mod.slice(mod.indexOf('async function setEventField'), mod.indexOf('async function setEventField') + 900);
if (!/raw === 'true'/.test(setter)) fail('is_public is not coerced from the select string to a boolean');
else pass('is_public is written as a boolean, not the string "true"');

// ---- buckets: the row banding and the state filter -------------------------
// eventBucket is pure, so it can be lifted out and exercised directly.
const bm = mod.match(/function eventBucket\(e\) \{[\s\S]*?\n\}/);
if (!bm) fail('eventBucket() is not defined');
else {
  const eventBucket = new Function('return (' + bm[0].replace('function eventBucket', 'function') + ')')();
  const cases = [
    [{ status: 'cancelled', stage: 'confirmed' }, 'cancelled', 'cancelled beats everything'],
    [{ status: 'completed' }, 'completed', 'completed status'],
    [{ status: 'announced', stage: 'wrapped' }, 'completed', 'a wrapped stage closes it out'],
    [{ status: 'planning', stage: 'reported' }, 'completed', 'reported closes it out'],
    [{ status: 'planning', stage: 'idea' }, 'potential', 'Blue Sky is potential'],
    [{ status: 'announced', stage: 'planning' }, 'active', 'announced is active'],
    [{ status: 'on_sale' }, 'active', 'on sale is active'],
    [{ status: 'sold_out' }, 'active', 'sold out is active'],
    [{ status: 'planning', stage: 'confirmed' }, 'active', 'a confirmed stage outranks a planning status'],
    [{ status: 'planning', stage: 'live' }, 'active', 'live is active'],
    [{ status: 'planning' }, 'potential', 'plain planning is potential'],
    [{}, 'potential', 'nothing set at all is potential'],
  ];
  let bad = 0;
  for (const [e, want, why] of cases) {
    const got = eventBucket(e);
    if (got !== want) { fail('eventBucket(' + JSON.stringify(e) + ') = ' + got + ', expected ' + want + ' - ' + why); bad++; }
  }
  if (!bad) pass('eventBucket buckets all ' + cases.length + ' cases correctly');

  // A future date must not demote an active event - the explicit ask.
  if (eventBucket({ status: 'announced', event_date: '2099-01-01' }) !== 'active') {
    fail('a future-dated announced event is not active');
  } else pass('future-dated confirmed events stay active');
}

// Every bucket needs a row style AND a chip style, or the chip row stops being a
// legend for the banding.
let styleGap = 0;
for (const k of ['active', 'potential', 'completed', 'cancelled']) {
  if (!src.includes('tr.ev-' + k)) { fail('no row banding for the ' + k + ' bucket'); styleGap++; }
  if (!src.includes('.ev-chip.bk-' + k)) { fail('no chip colour for the ' + k + ' bucket'); styleGap++; }
}
if (!styleGap) pass('all four buckets have matching row and chip colours');
if (!src.includes('border-left-color')) fail('banding is colour-only with no left edge');
else pass('banding carries a left edge, not colour alone');

// ---- multi-select ----------------------------------------------------------
if (/eventsDash = \{ rows: \[\], q: '', series: '',/.test(mod)) fail('filters are still single-value strings');
else if (!/series: \[\], status: \[\], year: \[\], bucket: \[\]/.test(mod)) fail('filter state is not four arrays');
else pass('series, status, year and state are all multi-select');

if (!mod.includes('const any = (arr, v) => !arr.length || arr.indexOf(v) >= 0;')) {
  fail('an empty filter list does not mean "no filter"');
} else pass('an empty filter list means no filter, not match-nothing');

if (!mod.includes('data-evnotdone')) fail('there is no "not completed" shortcut');
else if (!/eventsDash\.bucket = now === 'active,potential' \? \[\] : \['active', 'potential'\]/.test(mod)) {
  fail('"not completed" does not toggle active+potential');
} else pass('"not completed" toggles active + potential');

// The removed <select> handlers assigned a string to what is now an array.
let deadGap = 0;
for (const dead of ['data-evseries]', 'data-evstatus]', 'data-evyear]', 'data-evcompleted]']) {
  if (mod.includes(dead)) { fail('a handler for the removed ' + dead + ' filter is still wired up'); deadGap++; }
}
if (!deadGap) pass('no handler is left assigning a string to an array filter');

// Blue Sky has to be offered by name.
if (!/\['idea', 'Blue Sky'\]/.test(mod)) fail('the stage dropdown does not offer Blue Sky');
else pass('Blue Sky is offered as a stage');

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
