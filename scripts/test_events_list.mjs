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

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
