// Pull the money panel region out of dashboard.html and actually RUN the render
// path against fixture data. node --check proves the file parses; it cannot tell
// you that a function the renderer calls was deleted. This can.
import fs from 'node:fs';
const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

const START = 'const moneyState = { eventId: null, event: null };';
const END = '// Modal (used from the Events tab).';
const i = mod.indexOf(START), j = mod.indexOf(END, i);
if (i < 0 || j < 0) { console.error('REGION NOT FOUND'); process.exit(1); }
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
const sb = { from: () => ({ select: () => ({ eq: () => ({}) }) }) };
const expDash = { events: [] };
const confirm = () => false;
const prompt = () => null;
const document = { createElement: () => ({}), body: { appendChild() {} }, addEventListener() {} };

const fn = new Function(
  'escapeHtml,money,fmtDate,fmtNum,TODAY,toast,pickFrom,sbAll,sb,expDash,confirm,prompt,document',
  region + '\n; return { moneySectionsHTML, moneyFormHTML, moneyStatusHTML, moneyLinkHTML, moneyForm, moneyLink, moneyState, moneyFormDefs };'
);
const api = fn(escapeHtml, money, fmtDate, fmtNum, TODAY, toast, pickFrom, sbAll, sb, expDash, confirm, prompt, document);

api.moneyState.event = { id: 'e1', name: 'Come With 7-11', series: 'Come With Parties', event_date: '2026-07-11' };
api.moneyState.actors = [{ id: 'a1', display_name: 'Janelle Sochet' }];
api.moneyState.expCats = ['Talent', 'Venue', 'Marketing / Networking'];
api.moneyState.streams = ['DJ Gig fee', 'Door split'];

const data = {
  ev: { data: api.moneyState.event },
  tk: { data: [{ id: 't1', ticket_type: 'GA', quantity: 40, amount_paid: 800 }] },
  inc: { data: [
    { id: 'i1', amount: 500, category: 'DJ Gig fee', description: 'fee', date: '2026-07-11', status: 'invoiced', expected_amount: 500 },
    { id: 'i2', amount: 120, category: 'Bar', date: '2026-07-11', status: 'received', expected_amount: null },
  ] },
  exp: { data: [
    { id: 'x1', amount: 750, category: 'Talent', vendor: 'Test DJ', date: '2026-07-11', status: 'accrued', expected_amount: 750, due_date: '2026-08-01', cash_source: null },
    { id: 'x2', amount: 800, category: 'Venue', vendor: 'The Space', date: '2026-07-11', status: 'paid', expected_amount: 750, cash_source: 'bank' },
  ] },
  don: { data: [{ id: 'd1', amount: 50, donor_name: 'Anon' }] },
  spo: { data: [{ id: 's1', cash_amount: 500, in_kind_value: 100, tier: 'Gold', status: 'confirmed', actor: { display_name: 'Acme' } }] },
  sponsors: { data: [{ actor: { id: 'sp1', display_name: 'Acme' } }] },
};

let fails = 0;
const check = (label, f) => {
  try { const out = f();
    if (typeof out !== 'string' || !out.length) throw new Error('empty output');
    if (/undefined|\[object Object\]|NaN/.test(out)) throw new Error('renders literal: ' + (out.match(/undefined|\[object Object\]|NaN/) || [])[0]);
    console.log('PASS  ' + label + '  (' + out.length + ' chars)');
  } catch (e) { fails++; console.log('FAIL  ' + label + '  ' + e.message); }
};

check('panel, no form open', () => api.moneySectionsHTML(data));
for (const k of ['ticket', 'income', 'expense', 'donation', 'sponsorship']) {
  api.moneyForm.kind = k;
  check('panel with ' + k + ' form open', () => {
    const html = api.moneySectionsHTML(data);
    if (!html.includes('data-money-form-kind="' + k + '"')) throw new Error('form markup missing');
    return html;
  });
}
api.moneyForm.kind = null;

// The two questions the labels have to answer.
api.moneyForm.kind = 'expense';
const ex = api.moneySectionsHTML(data);
const need = ['Cost date', 'Due date', 'Category', 'Payee', 'Paid from', 'Has it been paid?'];
for (const label of need) {
  if (!ex.includes(label)) { fails++; console.log('FAIL  expense form is missing the "' + label + '" label'); }
}
if (!/value="2026-07-11"/.test(ex)) { fails++; console.log('FAIL  cost date does not default to the event date'); }
if (!ex.includes('data-status="paid"')) { fails++; console.log('FAIL  form does not start in the paid state'); }
console.log(fails ? '' : 'PASS  expense form labels both dates, defaults cost date to the event date');

// Category must be a controlled list, not a free-text box.
if (/<input[^>]*data-f="category"/.test(ex)) { fails++; console.log('FAIL  category is still a free-text input'); }
else if (!/<select[^>]*data-f="category"[^>]*>[\s\S]*?Talent/.test(ex)) { fails++; console.log('FAIL  category select is not populated'); }
else console.log('PASS  category is a select populated from the known set');

api.moneyForm.kind = null;
api.moneyLink.open = true;
api.moneyLink.rows = [{ id: 'z1', date: '2026-06-01', amount: 42, vendor: 'Old Charge', category: 'Software', event_id: null, status: 'paid' }];
check('link drawer open', () => api.moneySectionsHTML(data));

process.exit(fails ? 1 : 0);
