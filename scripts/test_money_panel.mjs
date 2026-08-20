// Pull the money panel region out of dashboard.html and actually RUN the render
// path against fixture data. node --check proves the file parses; it cannot tell
// you that a function the renderer calls was deleted. This can — and did, once.
//
//   node scripts/test_money_panel.mjs        (from the repo root)
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

const START = 'const moneyState = { eventId: null, event: null };';
const END = '// Modal (used from the Events tab).';
const i = mod.indexOf(START), j = mod.indexOf(END, i);
if (i < 0 || j < 0) { console.error('REGION NOT FOUND — markers moved'); process.exit(1); }
const region = mod.slice(i, j);

// Stubs for everything the region leans on from the wider module.
const escapeHtml = (v) => String(v == null ? '' : v)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const money = (n) => '$' + Number(n || 0).toFixed(2);
const fmtDate = (d) => String(d).slice(0, 10);
const fmtNum = (n) => String(n);
const TODAY = () => '2026-08-20';
const toast = () => {};
const pickFrom = async () => null;
const sbAll = async () => ({ data: [], error: null });
const sb = { from: () => ({ select: () => ({ eq: () => ({}) }) }), auth: { getSession: async () => ({ data: {} }) } };
const expDash = { events: [] };
const hub = { eventId: null, section: null };
const confirm = () => false;
const prompt = () => null;
const document = { createElement: () => ({}), body: { appendChild() {} }, addEventListener() {}, getElementById: () => null };

const fn = new Function(
  'escapeHtml,money,fmtDate,fmtNum,TODAY,toast,pickFrom,sbAll,sb,expDash,hub,confirm,prompt,document',
  region + '\n; return { moneySectionsHTML, moneyGridHTML, moneyLines, moneyRollup, moneyFormHTML,' +
  ' moneyForm, moneyGrid, moneyLink, moneyState };'
);
const api = fn(escapeHtml, money, fmtDate, fmtNum, TODAY, toast, pickFrom, sbAll, sb, expDash, hub, confirm, prompt, document);

api.moneyState.event = { id: 'e1', name: 'Come With 7-11', series: 'Come With Parties', event_date: '2026-07-11' };
api.moneyState.actors = [{ id: 'a1', display_name: 'Janelle Sochet' }];
api.moneyState.expCats = ['Talent', 'Venue', 'Marketing / Networking'];
api.moneyState.streams = ['DJ Gig fee', 'Door split'];

const data = {
  ev: { data: api.moneyState.event },
  tk: { data: [{ id: 't1', ticket_type: 'GA', quantity: 40, amount_paid: 800 }] },
  inc: { data: [
    { id: 'i1', amount: 500, category: 'DJ Gig fee', description: 'fee', date: '2026-07-11', status: 'invoiced', expected_amount: 500 },
    { id: 'i2', amount: 120, category: 'Door split', date: '2026-07-11', status: 'received', expected_amount: null },
  ] },
  exp: { data: [
    { id: 'x1', amount: 750, category: 'Talent', vendor: 'Test DJ', date: '2026-07-11', status: 'accrued', expected_amount: 750, due_date: '2026-08-01', cash_source: null },
    { id: 'x2', amount: 800, category: 'Venue', vendor: 'The Space', date: '2026-07-11', status: 'paid', expected_amount: 750, cash_source: 'bank' },
  ] },
  don: { data: [{ id: 'd1', amount: 50, donor_name: 'Anon' }] },
  spo: { data: [{ id: 's1', cash_amount: 500, in_kind_value: 100, tier: 'Gold', status: 'confirmed', actor: { display_name: 'Acme' } }] },
  sponsors: { data: [{ actor: { id: 'sp1', display_name: 'Acme' } }] },
  fc: { data: [
    { id: 'f1', category: 'Talent', direction: 'expense', planned_amount: 900, label: 'Headline DJ', confidence: 80, realized_at: null },
    { id: 'f2', category: 'Door split', direction: 'income', planned_amount: 400, label: 'Bar minimum', confidence: null, realized_at: null },
    { id: 'f3', category: 'Venue', direction: 'expense', planned_amount: 1000, label: 'Already booked', confidence: null, realized_at: '2026-08-01T00:00:00Z' },
  ] },
};

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);
const check = (label, f) => {
  try {
    const out = f();
    if (typeof out !== 'string' || !out.length) throw new Error('empty output');
    const bad = out.match(/undefined|\[object Object\]|NaN/);
    if (bad) throw new Error('renders literal "' + bad[0] + '"');
    pass(label + '  (' + out.length + ' chars)');
    return out;
  } catch (e) { fail(label + '  ' + e.message); return ''; }
};

// ---- the panel renders at all ------------------------------------------------
api.moneyGrid.open = {};
check('panel, all collapsed', () => api.moneySectionsHTML(data));

for (const k of ['ticket', 'income', 'expense', 'donation', 'sponsorship', 'forecast']) {
  api.moneyForm.kind = k;
  check('panel with the ' + k + ' form open', () => {
    const html = api.moneySectionsHTML(data);
    if (!html.includes('data-money-form-kind="' + k + '"')) throw new Error('form markup missing');
    return html;
  });
}
api.moneyForm.kind = null;

// ---- the expense form still answers Keith's two questions --------------------
api.moneyForm.kind = 'expense';
const ex = api.moneySectionsHTML(data);
for (const label of ['Cost date', 'Due date', 'Category', 'Payee', 'Paid from', 'Has it been paid?']) {
  if (!ex.includes(label)) fail('expense form is missing the "' + label + '" label');
}
if (!/value="2026-07-11"/.test(ex)) fail('cost date does not default to the event date');
if (/<input[^>]*data-f="category"/.test(ex)) fail('category is still a free-text input');
else pass('expense form: both dates labelled, cost date defaults to the event, category is a select');
api.moneyForm.kind = null;

// ---- the grid ----------------------------------------------------------------
const g = check('grid', () => api.moneyGridHTML(data));
if (!g.includes('pl-sec-revenue') || !g.includes('pl-sec-direct')) fail('grid is missing a section band');
if (!g.includes('data-mcat="cost|Talent"')) fail('grid has no Talent category row');
if (!/pl-caret">▸/.test(g)) fail('collapsed categories have no caret');
if (!g.includes('pl-net')) fail('grid has no NET line');
else pass('grid: section bands, category rows with carets, NET line');

// Rollup arithmetic is the thing most worth being sure about.
const roll = api.moneyRollup(api.moneyLines(data));
const T = (k) => roll[k] || { forecast: 0, booked: 0, settled: 0 };
const eq = (label, got, want) => Math.abs(got - want) < 0.005
  ? pass(label + ' = ' + money(want))
  : fail(label + ' = ' + money(got) + ', expected ' + money(want));
eq('Talent booked (accrued 750)', T('cost|Talent').booked, 750);
eq('Talent settled (nothing paid)', T('cost|Talent').settled, 0);
eq('Talent forecast (900 open)', T('cost|Talent').forecast, 900);
eq('Venue booked (800 paid)', T('cost|Venue').booked, 800);
eq('Venue settled', T('cost|Venue').settled, 800);
eq('Venue forecast EXCLUDES the realised line', T('cost|Venue').forecast, 0);
eq('Ticket sales booked', T('revenue|Ticket sales').booked, 800);
eq('DJ Gig fee booked (invoiced, not received)', T('revenue|DJ Gig fee').booked, 500);
eq('DJ Gig fee settled', T('revenue|DJ Gig fee').settled, 0);
eq('Door split forecast', T('revenue|Door split').forecast, 400);

// A realised forecast must not be counted twice.
if (g.includes('Already booked')) fail('a realised forecast line is still drawn in the grid');
else pass('realised forecast lines drop out of the forecast');

// Variance is only claimed where there is a plan to vary from.
if (!/\+\$?[\d,]*150\.00|\-\$150\.00|\$-150\.00/.test(g.replace(/<[^>]+>/g, ''))) {
  // Talent: booked 750 vs forecast 900 => -150
  fail('Talent variance (booked 750 vs forecast 900 = -150) is not shown');
} else pass('variance shown where a forecast exists');

// ---- expanded detail is editable --------------------------------------------
api.moneyGrid.open = { 'cost|Talent': true };
const opened = check('grid with Talent expanded', () => api.moneyGridHTML(data));
for (const need of ['data-medit="category"', 'data-medit="status"', 'data-medit="amount"', 'data-medit="date"']) {
  if (!opened.includes(need)) fail('expanded row is missing an inline editor: ' + need);
}
if (!opened.includes('data-mcommit="f1"')) fail('the forecast line has no Commit action');
if (!opened.includes('data-money-settle="expenses:x1"')) fail('the committed cost has no Pay action');
if (!opened.includes('m-forecast')) fail('forecast rows are not marked');
if (!fails) pass('expanded rows carry inline category/status/amount/date editors, Commit and Pay');

// Collapsed again, the detail must be gone.
api.moneyGrid.open = {};
if (api.moneyGridHTML(data).includes('data-medit="status"')) fail('collapsed categories still render detail rows');
else pass('collapsing hides the detail');

// The hub wraps the panel in its own header; repainting has to target the panel's
// own wrapper or that header disappears the first time a category is expanded.
if (!api.moneySectionsHTML(data).startsWith('<div data-money-panel>')) {
  fail('the panel has no [data-money-panel] wrapper for moneyRepaint to target');
} else pass('panel is wrapped so a repaint cannot eat the surrounding header');

// ---- the link drawer ---------------------------------------------------------
api.moneyLink.open = true;
api.moneyLink.rows = [{ id: 'z1', date: '2026-06-01', amount: 42, vendor: 'Old Charge', category: 'Software', event_id: null, status: 'paid' }];
check('link drawer open', () => api.moneySectionsHTML(data));
api.moneyLink.open = false;

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
