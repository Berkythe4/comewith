// Runs the Data Health renderer against fixtures. Same reasoning as
// test_money_panel.mjs: there is no local console for dashboard.html, and
// `node --check` proves the file parses, not that the renderer works.
//
//   node scripts/test_data_health.mjs        (from the repo root)
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

const START = 'const dhState = {';
const END = "$('panel-data-health').addEventListener";
const i = mod.indexOf(START), j = mod.indexOf(END, i);
if (i < 0 || j < 0) { console.error('REGION NOT FOUND — markers moved'); process.exit(1); }
const region = mod.slice(i, j);

const escapeHtml = (v) => String(v == null ? '' : v)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const fmtDT = (d) => String(d).slice(0, 16).replace('T', ' ');
const toast = () => {};
const sbAll = async () => ({ data: [], error: null });
const sb = { from: () => ({ select: () => ({}) }), rpc: async () => ({ data: {}, error: null }),
             auth: { getSession: async () => ({ data: {} }) } };
const activateTab = () => {};
const prompt = () => null;
const confirm = () => false;
const alert = () => {};

let painted = '';
const $ = () => ({ set innerHTML(v) { painted = v; }, get innerHTML() { return painted; } });

const fn = new Function(
  'escapeHtml,fmtDT,toast,sbAll,sb,activateTab,prompt,confirm,alert,$',
  region + '\n; return { renderDataHealth, dhState, dhCheckHTML, dhTrendHTML, dhRunsHTML, dhWaivedHTML };');
const api = fn(escapeHtml, fmtDT, toast, sbAll, sb, activateTab, prompt, confirm, alert, $);

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);

api.dhState.summary = [
  { category: 'finance', check_key: 'fin.expense_no_payee', severity: 'high', label: 'Cost with no payee linked', detail: 'blocks 1099', fix_hint: 'Link a payee', open_count: 7, waived_count: 1 },
  { category: 'finance', check_key: 'fin.donation_no_actor', severity: 'low', label: 'Donation with no donor actor', detail: 'repeat donors invisible', fix_hint: 'Link the donor', open_count: 8, waived_count: 0 },
  { category: 'events', check_key: 'ev.no_ticket_url', severity: 'high', label: 'Upcoming public event with no ticket link', detail: 'clicks cannot be backfilled', fix_hint: 'Set the ticket URL', open_count: 2, waived_count: 0 },
  { category: 'people', check_key: 'ppl.actor_no_role', severity: 'medium', label: 'Actor with no role', detail: 'hard to find', fix_hint: 'Assign a role', open_count: 1, waived_count: 0 },
  { category: 'content', check_key: 'ops.photo_no_credit', severity: 'low', label: 'Photo with no photographer', detail: 'no credit line', fix_hint: 'Credit in bulk', open_count: 502, waived_count: 0 },
];
api.dhState.rows = [];
for (let n = 0; n < 7; n++) api.dhState.rows.push({ check_key: 'fin.expense_no_payee', category: 'finance', severity: 'high', subject_table: 'expenses', subject_id: 'x' + n, subject_label: 'Food & Beverage · 300', fix_hint: 'Link a payee actor' });
for (let n = 0; n < 502; n++) api.dhState.rows.push({ check_key: 'ops.photo_no_credit', category: 'content', severity: 'low', subject_table: 'event_photos', subject_id: 'p' + n, subject_label: 'DI #1', fix_hint: 'Credit in bulk' });
api.dhState.rows.push({ check_key: 'ev.no_ticket_url', category: 'events', severity: 'high', subject_table: 'events', subject_id: 'e1', subject_label: 'Come With #2 · 2026-11-14', fix_hint: 'Set the ticket URL' });
api.dhState.rows.push({ check_key: 'ppl.actor_no_role', category: 'people', severity: 'medium', subject_table: 'actors', subject_id: 'a1', subject_label: 'Facebook', fix_hint: 'Assign a role' });
api.dhState.runs = [
  { id: 'r1', kind: 'audit', source: 'cron', total: 569, ran_at: '2026-08-20T07:00:00Z', by_severity: { high: 9, medium: 23, low: 537 }, summary: { 'ops.photo_no_credit': 502, 'ops.post_no_subject': 23 } },
  { id: 'r2', kind: 'autolink', source: 'cron', total: 12, ran_at: '2026-08-20T07:00:00Z', summary: { actor_roles_inferred: 11, guests_linked: 1, total: 12 } },
  { id: 'r3', kind: 'audit', source: 'migration-181', total: 580, ran_at: '2026-08-19T07:00:00Z', by_severity: {}, summary: {} },
];
api.dhState.waivers = [{ id: 'w1', check_key: 'fin.expense_no_payee', reason: 'Category, not a payee', waived_at: '2026-08-20T10:00:00Z' }];

const render = () => { api.renderDataHealth(); return painted; };

let html = '';
try { html = render(); } catch (e) { fail('renderDataHealth threw: ' + e.message); }
if (html) {
  const bad = html.match(/undefined|\[object Object\]|NaN/);
  if (bad) fail('renders literal "' + bad[0] + '"');
  else pass('panel renders (' + html.length + ' chars)');
  for (const need of ['Data Health', 'data-dh-run', 'data-dh-link="dry"', 'data-dh-link="apply"']) {
    if (!html.includes(need)) fail('toolbar is missing ' + need);
  }
  if (!html.includes('Finance') || !html.includes('Events') || !html.includes('People')) fail('category bands missing');
  else pass('toolbar (run / preview / apply) and category bands present');
  // The counts on the cards must come from the ROWS, not the summary, or a stale
  // summary would quietly disagree with the list underneath it.
  if (!/Needs attention[\s\S]{0,120}>8</.test(html)) fail('high-severity card should read 8 (7 payees + 1 ticket url)');
  else pass('severity cards count the actual findings');
}

// Collapsed by default; expanding shows rows with a fix link and a waive button.
if (html.includes('data-dh-waive')) fail('checks are expanded before being clicked');
else pass('checks start collapsed');

api.dhState.open['fin.expense_no_payee'] = true;
const open = render();
if (!open.includes('data-dh-waive')) fail('expanding a check shows no waive action');
if (!open.includes('data-dh-goto="expenses"')) fail('expanding a check shows no deep link to the fixing screen');
else pass('expanded findings carry a deep link and a waive action');

// A cap must never read as "that is all there is" (LEARNINGS §18).
api.dhState.open = { 'ops.photo_no_credit': true };
const capped = render();
if (!/Showing 40 of 502/.test(capped)) fail('the 502-row check does not say how many it is hiding');
else pass('a capped list says what it is not showing');

// Trend compares AUDIT runs only — comparing against an autolink row would read
// as a 557-finding improvement that never happened.
api.dhState.open = {};
const trend = api.dhTrendHTML();
if (!/11/.test(trend)) fail('trend should read 11 (569 vs the previous audit at 580), got: ' + trend.replace(/<[^>]+>/g, ''));
else pass('trend compares audit runs only: down 11');

// Waivers are hidden until asked for, and always show their reason.
if (render().includes('Category, not a payee')) fail('waived items are shown before being asked for');
api.dhState.showWaived = true;
if (!render().includes('Category, not a payee')) fail('waived list does not show the reason');
else pass('waivers hidden by default, reason shown when opened');

// The run history is the validation record — it has to say what each run DID.
const runs = api.dhRunsHTML();
if (!/actor roles inferred 11/.test(runs)) fail('run history does not summarise what the auto-link changed');
else pass('run history summarises every run');

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
