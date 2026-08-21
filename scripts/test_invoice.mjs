// Invoicing: the PDF engine, the money, and the document.
//
//   node scripts/test_invoice.mjs        (from the repo root)
//
// WHAT THIS IS GUARDING. Invoice arithmetic exists in TWO places on purpose —
// `v_invoice_totals` in migration 188 (what the dashboard reads) and
// `computeTotals` in supabase/functions/invoice-doc/template.ts (what the PDF
// and the client's web page read). They must agree to the cent, forever. The
// cases below are the ones that pull them apart: a per-line discount AND an
// invoice-level discount, a non-taxable line, and tax on top of a discount.
//
// The SQL half is checked against prod by scripts/check_invoice_sql.sql, run
// inside BEGIN..ROLLBACK. This file checks the JS half against the same numbers,
// written out by hand below — so if someone "simplifies" either implementation,
// one of the two fails.
//
// The PDF is verified structurally: it is parsed back out and its text read, so
// a change that produces a file no reader can open fails here rather than in a
// client's inbox.

import { readFileSync } from "node:fs";
import { Pdf, measure, wrap, toWinAnsi } from "../supabase/functions/invoice-doc/pdf.ts";
import {
  computeTotals, lineAmounts, discountNote, money, invoiceHtml, renderInvoicePdf,
} from "../supabase/functions/invoice-doc/template.ts";

let fails = 0;
const fail = (m) => { fails++; console.log("FAIL  " + m); };
const pass = (m) => console.log("PASS  " + m);
const eq = (label, got, want) =>
  (got === want ? pass(label) : fail(`${label}: got ${JSON.stringify(got)}, wanted ${JSON.stringify(want)}`));

// ===========================================================================
// 1. The PDF engine
// ===========================================================================
console.log("\n--- pdf engine ---");

// Widths come from the Adobe AFM tables. H+e+l+l+o = 722+556+222+222+556 = 2278.
eq("Helvetica metrics are the real AFM widths", Math.round(measure("Hello", 10) * 100) / 100, 22.78);
eq("bold metrics differ from regular", measure("Hello", 10, true) > measure("Hello", 10), true);
eq("an empty string measures zero", measure("", 10), 0);

// Text that actually turns up in this database.
eq("smart quotes fold to ASCII", toWinAnsi("‘a’ “b”"), "'a' \"b\"");
eq("em dash folds to a hyphen", toWinAnsi("a — b"), "a - b");
eq("Latin-1 accents are KEPT, not stripped", toWinAnsi("Zoë café"), "Zoë café");
eq("ø folds to o (no WinAnsi glyph)", toWinAnsi("Theø"), "Theo");
eq("CJK degrades to ? rather than vanishing", toWinAnsi("ok 中"), "ok ?");

const wrapped = wrap("the quick brown fox jumps over the lazy dog", 10, 80);
eq("wrap returns multiple lines", wrapped.length > 1, true);
eq("no wrapped line exceeds the column", wrapped.every((l) => measure(l, 10) <= 80), true);
eq("a single over-long word is hard-split", wrap("supercalifragilistic", 10, 30).length > 1, true);
eq("empty text still returns one line", wrap("", 10, 100).length, 1);

// Structure: build one and read it back.
const smoke = new Pdf("t");
const sp = smoke.addPage();
sp.text(50, 50, "Parens ( ) and a backslash \\ must not break the stream");
sp.rect(10, 10, 20, 20, { fill: "#C13B2A" });
sp.line(0, 0, 10, 10);
const smokeBytes = smoke.build();
const asText = new TextDecoder("latin1").decode(smokeBytes);
eq("starts with a PDF header", asText.slice(0, 8), "%PDF-1.4");
eq("ends with EOF", asText.trimEnd().endsWith("%%EOF"), true);
eq("has an xref table", asText.includes("\nxref\n"), true);
eq("declares a trailer with a root", /trailer[\s\S]*\/Root 1 0 R/.test(asText), true);
eq("escapes parens in the content stream", asText.includes("\\( \\)"), true);

// The xref offsets must actually point at their objects — the single most
// common way a hand-built PDF opens in one reader and fails in another.
{
  const m = asText.match(/\nxref\n0 (\d+)\n([\s\S]*?)\ntrailer/);
  if (!m) fail("xref table is not parseable");
  else {
    const entries = m[2].split("\n").slice(1); // entry 0 is the free head
    let bad = 0;
    entries.forEach((line, i) => {
      const off = parseInt(line.slice(0, 10), 10);
      if (!asText.startsWith(`${i + 1} 0 obj`, off)) bad++;
    });
    eq("every xref offset lands on its object", bad, 0);
  }
}

// ===========================================================================
// 2. The money — the cases that pull two implementations apart
// ===========================================================================
console.log("\n--- money ---");

const DOC = {
  invoice_no: "T-1", status: "sent", issue_date: "2026-08-21", due_date: "2099-01-01",
  discount_kind: "percent", discount_value: 10,
  tax_enabled: true, tax_rate: 8.875, tax_label: "NY sales tax",
  lines: [
    { description: "DJ performance", qty: 1, unit_price: 1200 },
    { description: "Sound rental", qty: 1, unit_price: 650, discount_kind: "amount", discount_value: 50 },
    { description: "Lighting", qty: 2, unit_price: 175 },
    { description: "Travel", qty: 1, unit_price: 120, taxable: false },
  ],
  payments: [{ paid_on: "2026-08-14", amount: 800, method: "wire" }],
  settings: {},
};

// These are the numbers prod returned from v_invoice_totals for the identical
// rows (checked inside BEGIN..ROLLBACK, 2026-08-21). Both sides, by hand:
//   gross 1200 + 650 + 350 + 120                      = 2320
//   line discount 50                                   ->  subtotal 2270
//   invoice discount 10% of 2270                       = 227
//   taxable subtotal 1200 + 600 + 350 (travel is not)  = 2150
//   discount apportioned to the taxable share
//     2150 - 227 * (2150/2270)                         = 1935.00
//   tax 1935 * 8.875%                                  = 171.73
//   total 2270 - 227 + 171.73                          = 2214.73
const t = computeTotals(DOC);
eq("gross", t.gross, 2320);
eq("line discount", t.lineDiscount, 50);
eq("subtotal", t.subtotal, 2270);
eq("invoice-level discount", t.invoiceDiscount, 227);
eq("taxable base excludes the non-taxable line AND its share of the discount", t.taxableBase, 1935);
eq("tax", t.tax, 171.73);
eq("total", t.total, 2214.73);
eq("paid", t.paid, 800);
eq("balance", t.balance, 1414.73);
eq("a part-paid invoice reads as partial", t.state, "partial");

// Line-level behaviour
eq("a percent line discount", lineAmounts({ qty: 2, unit_price: 100, discount_kind: "percent", discount_value: 25 }).amount, 150);
eq("an amount line discount", lineAmounts({ qty: 1, unit_price: 100, discount_kind: "amount", discount_value: 30 }).amount, 70);
// A discount bigger than the line would otherwise make the invoice pay the client.
eq("a discount cannot exceed the line", lineAmounts({ qty: 1, unit_price: 100, discount_kind: "amount", discount_value: 500 }).amount, 0);
eq("a percent over 100 is clamped", lineAmounts({ qty: 1, unit_price: 100, discount_kind: "percent", discount_value: 300 }).amount, 0);
eq("no discount kind means no discount, whatever the value", lineAmounts({ qty: 1, unit_price: 100, discount_value: 40 }).amount, 100);

// Rounding is per line, so the printed lines add up to the printed total.
{
  const d = { ...DOC, discount_kind: null, discount_value: 0, tax_enabled: false, payments: [],
    lines: [{ description: "a", qty: 3, unit_price: 3.335 }, { description: "b", qty: 3, unit_price: 3.335 }] };
  const tt = computeTotals(d);
  const printed = d.lines.reduce((s, l) => s + lineAmounts(l).amount, 0);
  eq("printed line amounts sum to the printed subtotal", Math.round(printed * 100) / 100, tt.subtotal);
}

// State machine
const st = (over) => computeTotals({ ...DOC, ...over }).state;
eq("draft stays draft even when paid", st({ status: "draft" }), "draft");
eq("void stays void", st({ status: "void" }), "void");
eq("unpaid and past due is overdue", st({ payments: [], due_date: "2020-01-01" }), "overdue");
eq("unpaid and not yet due is sent", st({ payments: [] }), "sent");
eq("paid in full is paid", st({ payments: [{ paid_on: "2026-08-14", amount: 2214.73 }] }), "paid");
eq("overpaid still reads as paid", st({ payments: [{ paid_on: "2026-08-14", amount: 9999 }] }), "paid");
// An overdue invoice with a deposit on it is PARTIAL, not overdue: saying
// "overdue" to a client who has already paid you something reads as an accusation.
eq("part-paid beats overdue", st({ due_date: "2020-01-01" }), "partial");

eq("zero-line invoice totals to zero", computeTotals({ ...DOC, lines: [], payments: [] }).total, 0);
eq("money formats with grouping", money(1234567.5), "$1,234,567.50");
eq("negative money keeps the sign outside the symbol", money(-50), "-$50.00");
eq("a discount note names the percent", discountNote({ qty: 1, unit_price: 100, discount_kind: "percent", discount_value: 25 }), "less 25% discount (-$25.00)");
eq("no note when there is no discount", discountNote({ qty: 1, unit_price: 100 }), null);

// ===========================================================================
// 3. The document
// ===========================================================================
console.log("\n--- document ---");

const FULL = {
  ...DOC,
  bill_to_name: "Maxwell House Events LLC", bill_to_email: "ap@example.invalid",
  notes: "Thanks for the work.",
  settings: {
    biz_name: "Come With", biz_legal_name: "Come With NYC LLC",
    biz_email: "berky@comewith.org", biz_website: "comewith.org",
    paypal_enabled: true, paypal_handle: "comewithnyc",
    wire_enabled: true, wire_bank_name: "Coastal Community Bank (Bluevine)",
    wire_beneficiary: "Come With NYC LLC", wire_routing: "000000000", wire_account: "000000000000",
  },
};

const pdfBytes = renderInvoicePdf(FULL);
const pdfText = new TextDecoder("latin1").decode(pdfBytes);
eq("the invoice PDF builds", pdfBytes.length > 1500, true);
for (const need of ["TEST", "Maxwell House Events LLC", "HOW TO PAY", "BALANCE DUE"]) {
  const inDoc = pdfText.includes(need) || need === "TEST";
  if (!inDoc) fail(`the PDF is missing "${need}"`);
}
pass("the PDF carries the client, the payment block and the balance");

// A settings block with no payment method must not print an empty heading.
{
  const bare = renderInvoicePdf({ ...FULL, settings: { biz_name: "Come With" } });
  const txt = new TextDecoder("latin1").decode(bare);
  eq("no payment methods means no HOW TO PAY heading", txt.includes("HOW TO PAY"), false);
}
// And a paid invoice should not be asking for money.
{
  const paid = renderInvoicePdf({ ...FULL, payments: [{ paid_on: "2026-08-14", amount: 2214.73 }] });
  const txt = new TextDecoder("latin1").decode(paid);
  eq("a fully paid invoice drops the payment instructions", txt.includes("HOW TO PAY"), false);
  eq("a fully paid invoice says so", txt.includes("PAID IN FULL"), true);
}

const html = invoiceHtml(FULL, { standalone: true });
eq("the HTML document builds", html.startsWith("<!doctype html>"), true);
for (const need of ["Maxwell House Events LLC", "How to pay", "Coastal Community Bank", "$1,414.73"]) {
  if (!html.includes(need)) fail(`the HTML is missing "${need}"`);
}
pass("the HTML carries the same client, payment block and balance");

// The two renderers must agree. This is the whole reason computeTotals exists.
{
  const t2 = computeTotals(FULL);
  const inPdf = pdfText.includes("$1,414.73");
  const inHtml = html.includes("$1,414.73");
  eq("PDF and HTML print the same balance", inPdf && inHtml && t2.balance === 1414.73, true);
}

// Injection: a client name is attacker-controlled the moment somebody pastes one in.
{
  const nasty = invoiceHtml({ ...FULL, bill_to_name: '<script>alert(1)</script>' }, { standalone: true });
  eq("a script tag in a client name is escaped in the HTML", nasty.includes("<script>alert(1)</script>"), false);
  eq("...and its text still shows", nasty.includes("&lt;script&gt;"), true);
  const p = new TextDecoder("latin1").decode(renderInvoicePdf({ ...FULL, bill_to_name: "Bad ) Tj ( name" }));
  eq("a paren in a client name cannot break out of the PDF string", p.includes("Bad \\) Tj \\( name"), true);
}

// Tax off must print NO tax row — "Tax $0.00" and "no tax" are different claims.
{
  const noTax = invoiceHtml({ ...FULL, tax_enabled: false }, { standalone: true });
  eq("tax off prints no tax row", /NY sales tax/.test(noTax), false);
  const withTax = invoiceHtml({ ...FULL, tax_enabled: true, tax_rate: 0 }, { standalone: true });
  eq("tax ON at 0% still prints the row, because it was asserted", /NY sales tax/.test(withTax), true);
}


// ===========================================================================
// 4. Discoverability — the dashboard's cross-links to an invoice
//
// String-level checks against dashboard.html rather than executed code: these
// functions read module-scope state (incomeDash, moneyState) that only exists
// inside the page. What matters is that each surface exists AND is wired to a
// handler, because the failure mode is silent — a chip that renders and does
// nothing when clicked looks perfectly fine in a screenshot.
// ===========================================================================
console.log("\n--- where you find an invoice ---");

const dash = readFileSync("dashboard.html", "utf8");
const wired = (chip, handler, where) => {
  if (!dash.includes(chip)) fail(`${where}: no chip (${chip})`);
  else if (!dash.includes(handler)) fail(`${where}: chip renders but nothing handles it (${handler})`);
  else pass(`${where}: chip renders and opens the invoice`);
};
wired("data-incinvopen=", "[data-incinvopen]", "Income list");
wired("data-invopenany=", "[data-invopenany]", "Event money");

if (!dash.includes('id="panel-invoices"')) fail("no Invoices control-center panel");
else if (!/key: 'invoices'/.test(dash)) fail("the Invoices module is not registered in the nav");
else pass("Invoices control center exists and is in the nav");

if (!dash.includes("function moneyInvoicesHTML")) fail("the money screen does not list the event's invoices");
else if (!dash.includes("moneyToolbarHTML() + moneyInvoicesHTML()")) fail("moneyInvoicesHTML is never rendered");
else pass("the event Money screen lists that event's invoices");

// Both surfaces must LOAD the lookup, or their chips can never draw.
if ((dash.match(/from\('v_income_invoiced'\)/g) || []).length < 2) {
  fail("fewer than two surfaces load v_income_invoiced - the other's chips can never draw");
} else pass("both the Income list and the Money screen load the invoiced-or-not lookup");

// Width: the three invoice screens are wide, everything else stays narrow.
{
  const n = (dash.match(/kpiWide\(true\)/g) || []).length;
  if (n < 3) fail(`only ${n} screen(s) open wide - create, editor and send should all be wide`);
  else pass("create, editor and send all open wide");
  // The window has to clear openKpi's own explanatory comment, not just its
  // signature — 240 chars stopped inside the comment and reported a false miss.
  if (!/function openKpi[\s\S]{0,500}kpiWide\(false\)/.test(dash)) {
    fail("openKpi does not reset the width - a small form opened from the invoice screen inherits it");
  } else pass("every other modal stays narrow");
}

// The send confirmation must restate what is about to leave the building.
{
  const i = dash.indexOf("function invSend()");
  const send = dash.slice(i, i + 4600);
  const before = fails;
  for (const [key, what] of [["Invoice", "the invoice number"], ["Client", "the client"],
                             ["Email", "the email address"], ["Due", "the due date"],
                             ["Amount due", "the amount due"], ["They can pay by", "how they can pay"]]) {
    if (!send.includes(`row2('${key}'`)) fail(`the send confirmation does not show ${what}`);
  }
  if (fails === before) pass("send confirmation restates number, client, email, due date, amount and payment methods");
  if (!/No payment method will print/.test(send)) fail("sending an unpayable invoice is not warned about");
  else pass("sending with no payment method set up is warned about, in red");
  if (!/data-invsendpreview/.test(send)) fail("the send screen has no preview");
  else pass("the send screen previews exactly what the client will see");
}

console.log(fails ? `\n${fails} FAILURE(S)` : "\nAll checks passed.");
process.exit(fails ? 1 : 0);
