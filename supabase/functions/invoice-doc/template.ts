// template.ts — the invoice itself: one layout, rendered to PDF and to HTML.
//
// BOTH RENDERERS READ THE SAME `InvoiceDoc`. That is the whole point of this
// file. The client gets a PDF attached to an email and a web page behind a link,
// and if those two were built from different code they would eventually disagree
// about a total — which is the one thing an invoice may never do. Money is
// computed ONCE, in `computeTotals`, and both renderers only draw what it
// returns.
//
// The numbers here mirror `v_invoice_totals` in migration 188 line for line,
// including the proportional split of an invoice-level discount across the
// taxable base. The view is the authority for the dashboard; this is the
// authority for the document; `scripts/test_invoice.mjs` checks them against
// each other so a change to one that is not made to the other fails a test
// rather than a client's arithmetic.
//
// TYPOGRAPHY NOTE. The site's display face is Bebas Neue, which cannot be used
// here — embedding a font would mean shipping the file and the licence into an
// edge function. The header instead uses Helvetica-Bold tracked out wide, which
// reads as the same kind of mark at a glance. The HTML version, which runs in a
// browser, uses the real Bebas Neue.

import { Pdf, PAGE, measure, wrap } from "./pdf.ts";

// ---- brand -----------------------------------------------------------------
export const INK = "#1A1410";
export const CREAM = "#F2EDE6";
export const RED = "#C13B2A";
export const MID = "#8A7F72";
export const RULE = "#DDD5C9";
export const GREEN = "#3B6D11";

// ---- shapes ----------------------------------------------------------------
export type Line = {
  description: string;
  detail?: string | null;
  qty: number;
  unit_price: number;
  discount_kind?: "amount" | "percent" | null;
  discount_value?: number | null;
  taxable?: boolean;
};
export type Payment = {
  paid_on: string;
  amount: number;
  method?: string | null;
  reference?: string | null;
};
export type Settings = {
  biz_name?: string | null;
  biz_legal_name?: string | null;
  biz_address?: string | null;
  biz_email?: string | null;
  biz_phone?: string | null;
  biz_website?: string | null;
  tax_id?: string | null;
  paypal_enabled?: boolean;
  paypal_handle?: string | null;
  paypal_note?: string | null;
  wire_enabled?: boolean;
  wire_bank_name?: string | null;
  wire_beneficiary?: string | null;
  wire_routing?: string | null;
  wire_account?: string | null;
  wire_swift?: string | null;
  wire_bank_address?: string | null;
  wire_note?: string | null;
  // Anything else you take money by - Venmo, Zelle, Cash App, Wise. Order is
  // display order; presence is what enables it.
  extra_methods?: Array<{ label?: string; detail?: string; note?: string }> | null;
  footer_note?: string | null;
};
export type InvoiceDoc = {
  invoice_no: string;
  status: string;
  issue_date: string;
  due_date?: string | null;
  bill_to_name?: string | null;
  bill_to_email?: string | null;
  bill_to_address?: string | null;
  event_name?: string | null;
  event_date?: string | null;
  currency?: string;
  discount_kind?: "amount" | "percent" | null;
  discount_value?: number | null;
  tax_enabled?: boolean;
  tax_rate?: number | null;
  tax_label?: string | null;
  notes?: string | null;
  terms_text?: string | null;
  pay_paypal?: boolean;
  pay_wire?: boolean;
  pay_extra?: boolean;
  lines: Line[];
  payments?: Payment[];
  settings: Settings;
  pay_url?: string | null;
};

// ---- money -----------------------------------------------------------------
const r2 = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;
export const money = (n: number, cur = "USD") => {
  const sign = n < 0 ? "-" : "";
  const v = Math.abs(r2(n)).toFixed(2);
  const [whole, frac] = v.split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return `${sign}${cur === "USD" ? "$" : ""}${grouped}.${frac}`;
};

export function lineAmounts(l: Line) {
  const gross = r2((Number(l.qty) || 0) * (Number(l.unit_price) || 0));
  const dv = Number(l.discount_value) || 0;
  let disc = 0;
  if (l.discount_kind === "percent") disc = r2((gross * Math.min(dv, 100)) / 100);
  else if (l.discount_kind === "amount") disc = r2(Math.min(dv, gross));
  return { gross, discount: disc, amount: r2(gross - disc) };
}

export function computeTotals(doc: InvoiceDoc) {
  let gross = 0, lineDiscount = 0, subtotal = 0, taxableSubtotal = 0;
  for (const l of doc.lines || []) {
    const a = lineAmounts(l);
    gross += a.gross;
    lineDiscount += a.discount;
    subtotal += a.amount;
    if (l.taxable !== false) taxableSubtotal += a.amount;
  }
  gross = r2(gross); lineDiscount = r2(lineDiscount);
  subtotal = r2(subtotal); taxableSubtotal = r2(taxableSubtotal);

  const dv = Number(doc.discount_value) || 0;
  let invoiceDiscount = 0;
  if (doc.discount_kind === "percent") invoiceDiscount = r2((subtotal * Math.min(dv, 100)) / 100);
  else if (doc.discount_kind === "amount") invoiceDiscount = r2(Math.min(dv, subtotal));

  // The invoice-level discount is apportioned to the taxable part in the same
  // ratio it bears to the whole, so turning tax on never charges tax on money
  // the client is not paying.
  const share = subtotal > 0 ? taxableSubtotal / subtotal : 0;
  const taxableBase = r2(Math.max(taxableSubtotal - invoiceDiscount * share, 0));
  const tax = doc.tax_enabled ? r2((taxableBase * (Number(doc.tax_rate) || 0)) / 100) : 0;
  const total = r2(subtotal - invoiceDiscount + tax);
  const paid = r2((doc.payments || []).reduce((t, p) => t + (Number(p.amount) || 0), 0));
  const balance = r2(total - paid);

  let state = doc.status;
  if (doc.status !== "draft" && doc.status !== "void") {
    if (paid > 0 && paid >= total) state = "paid";
    else if (paid > 0) state = "partial";
    else if (doc.due_date && doc.due_date < new Date().toISOString().slice(0, 10)) state = "overdue";
    else state = "sent";
  }
  return { gross, lineDiscount, subtotal, invoiceDiscount, taxableBase, tax, total, paid, balance, state };
}

export const STATE_LABEL: Record<string, string> = {
  draft: "DRAFT", sent: "SENT", partial: "PARTIALLY PAID",
  paid: "PAID", overdue: "OVERDUE", void: "VOID",
};
const STATE_COLOR: Record<string, string> = {
  draft: MID, sent: INK, partial: "#854F0B", paid: GREEN, overdue: RED, void: MID,
};

const fmtDate = (d?: string | null) => {
  if (!d) return "";
  const s = String(d).slice(0, 10);
  const [y, m, day] = s.split("-").map(Number);
  if (!y || !m || !day) return s;
  const M = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${M[m - 1]} ${day}, ${y}`;
};

/** The discount note that prints under a discounted line. */
export function discountNote(l: Line): string | null {
  const dv = Number(l.discount_value) || 0;
  if (!l.discount_kind || dv <= 0) return null;
  const a = lineAmounts(l);
  return l.discount_kind === "percent"
    ? `less ${dv}% discount (${money(-a.discount)})`
    : `less discount ${money(-a.discount)}`;
}

// ===========================================================================
// PDF
// ===========================================================================
const M = 48;                    // page margin
const COL_QTY = 330, COL_RATE = 400, COL_AMT = 480;
const COL_W = 84;                // width of each right-aligned money column

export function renderInvoicePdf(doc: InvoiceDoc): Uint8Array {
  const t = computeTotals(doc);
  const s = doc.settings || {};
  const cur = doc.currency || "USD";
  const pdf = new Pdf(`Invoice ${doc.invoice_no}`, s.biz_name || "Come With");

  let page = pdf.addPage();
  let y = 0;

  // withHead is only true when the LINE TABLE is what continues. Drawing the
  // column headers on a page that carries the totals or the payment block
  // produced a "DESCRIPTION QTY RATE AMOUNT" header with nothing underneath it.
  const newPage = (withHead: boolean) => {
    page = pdf.addPage();
    y = M;
    // Continuation header, so a page 2 that arrives on its own is identifiable.
    page.text(M, y, `${s.biz_name || "Come With"} — Invoice ${doc.invoice_no} (continued)`,
      { size: 8, color: MID });
    y += 22;
    if (withHead) tableHead();
  };
  const room = (need: number, withHead = false) => {
    if (y + need > PAGE.H - 52) { newPage(withHead); return true; }
    return false;
  };

  // ---- masthead ----
  page.rect(0, 0, PAGE.W, 104, { fill: INK });
  page.text(M, 40, (s.biz_name || "COME WITH").toUpperCase(),
    { size: 24, bold: true, color: CREAM, spacing: 2.4 });
  const sub = [s.biz_website, s.biz_email].filter(Boolean).join("  ·  ");
  if (sub) page.text(M, 62, sub, { size: 8, color: "#B9AE9E", spacing: 0.8 });
  page.text(PAGE.W - M - 200, 38, "INVOICE",
    { size: 22, bold: true, color: CREAM, align: "right", width: 200, spacing: 3 });
  page.text(PAGE.W - M - 200, 60, doc.invoice_no,
    { size: 10, color: RED, align: "right", width: 200, spacing: 1.2 });
  page.rect(0, 104, PAGE.W, 3, { fill: RED });

  y = 132;

  // ---- from / bill-to / meta ----
  const fromLines = [
    s.biz_legal_name || s.biz_name || "Come With",
    ...(s.biz_address || "").split("\n"),
    s.biz_email || "", s.biz_phone || "",
    s.tax_id ? `Tax ID ${s.tax_id}` : "",
  ].filter((x) => x && x.trim());
  const toLines = [
    doc.bill_to_name || "—",
    ...(doc.bill_to_address || "").split("\n"),
    doc.bill_to_email || "",
  ].filter((x) => x && x.trim());

  page.text(M, y, "FROM", { size: 7, color: MID, spacing: 1.4 });
  page.text(M + 190, y, "BILL TO", { size: 7, color: MID, spacing: 1.4 });
  let fy = y + 15, ty = y + 15;
  fromLines.forEach((l, i) => { page.text(M, fy, l, { size: 9, bold: i === 0 }); fy += 13; });
  toLines.forEach((l, i) => { page.text(M + 190, ty, l, { size: 9, bold: i === 0 }); ty += 13; });

  // Right column: the three dates and the amount due, which is the number the
  // client is actually looking for.
  const rx = PAGE.W - M - 190;
  const meta: [string, string][] = [
    ["Issued", fmtDate(doc.issue_date)],
    ["Due", doc.due_date ? fmtDate(doc.due_date) : "On receipt"],
  ];
  if (doc.event_name) meta.push(["Event", doc.event_name + (doc.event_date ? ` · ${fmtDate(doc.event_date)}` : "")]);
  let my = y;
  meta.forEach(([k, v]) => {
    page.text(rx, my, k.toUpperCase(), { size: 7, color: MID, spacing: 1.2 });
    page.text(rx, my + 12, v, { size: 9 });
    my += 30;
  });

  y = Math.max(fy, ty, my) + 10;

  // Amount-due badge
  const badgeH = 46;
  page.rect(rx - 10, y, 190 + 10, badgeH, { fill: "#F7F3EC" });
  page.text(rx, y + 15, t.balance > 0 ? "AMOUNT DUE" : "AMOUNT", { size: 7, color: MID, spacing: 1.4 });
  page.text(rx, y + 36, money(t.balance > 0 ? t.balance : t.total, cur),
    { size: 17, bold: true, color: t.state === "overdue" ? RED : INK });
  page.text(rx + 190 - 2, y + 15, STATE_LABEL[t.state] || t.state.toUpperCase(),
    { size: 7, bold: true, color: STATE_COLOR[t.state] || INK, align: "right", width: 0, spacing: 1 });

  y += badgeH + 12;

  // ---- line table ----
  function tableHead() {
    page.text(M, y, "DESCRIPTION", { size: 7, color: MID, spacing: 1.2 });
    page.text(COL_QTY, y, "QTY", { size: 7, color: MID, align: "right", width: 40, spacing: 1.2 });
    page.text(COL_RATE, y, "RATE", { size: 7, color: MID, align: "right", width: COL_W - 20, spacing: 1.2 });
    page.text(COL_AMT, y, "AMOUNT", { size: 7, color: MID, align: "right", width: COL_W, spacing: 1.2 });
    y += 8;
    page.line(M, y, PAGE.W - M, y, { color: INK, w: 0.9 });
    y += 14;
  }
  tableHead();

  const descW = COL_QTY - M - 55;
  for (const l of doc.lines || []) {
    const a = lineAmounts(l);
    const dLines = wrap(l.description || "", 9.5, descW, true);
    const detail = l.detail ? wrap(l.detail, 8.5, descW) : [];
    const note = discountNote(l);
    room(dLines.length * 13 + detail.length * 11 + (note ? 11 : 0) + 12, true);

    const top = y;
    dLines.forEach((ln) => { page.text(M, y, ln, { size: 9.5, bold: true }); y += 13; });
    detail.forEach((ln) => { page.text(M + 8, y, ln, { size: 8.5, color: MID }); y += 11; });
    if (note) { page.text(M + 8, y, note, { size: 8.5, color: RED }); y += 11; }

    // Numbers sit on the first line of the description, not the last.
    const qty = Number(l.qty) || 0;
    page.text(COL_QTY, top, qty % 1 === 0 ? String(qty) : String(qty),
      { size: 9.5, align: "right", width: 40 });
    page.text(COL_RATE, top, money(Number(l.unit_price) || 0, cur),
      { size: 9.5, align: "right", width: COL_W - 20 });
    page.text(COL_AMT, top, money(a.amount, cur), { size: 9.5, align: "right", width: COL_W });

    y += 5;
    page.line(M, y, PAGE.W - M, y, { color: RULE, w: 0.5 });
    y += 10;
  }
  if (!(doc.lines || []).length) {
    page.text(M, y, "No line items.", { size: 9, color: MID });
    y += 20;
  }

  // ---- totals ----
  room(150);
  const tx = COL_RATE - 60;
  const tw = PAGE.W - M - tx;
  const row = (label: string, value: string, o: { bold?: boolean; color?: string; size?: number } = {}) => {
    page.text(tx, y, label, { size: o.size ?? 9, color: o.color ?? MID, bold: o.bold });
    page.text(tx, y, value, { size: o.size ?? 9, align: "right", width: tw, bold: o.bold, color: o.color ?? INK });
    y += 14;
  };
  y += 2;
  row("Subtotal", money(t.subtotal, cur));
  if (t.lineDiscount > 0) row("Line discounts", money(-t.lineDiscount, cur), { color: MID });
  if (t.invoiceDiscount > 0) {
    const lbl = doc.discount_kind === "percent" ? `Discount (${Number(doc.discount_value)}%)` : "Discount";
    row(lbl, money(-t.invoiceDiscount, cur), { color: RED });
  }
  // A zero tax row is only shown when tax is deliberately ON. "Tax $0.00" and
  // "no tax line" are different claims.
  if (doc.tax_enabled) {
    row(`${doc.tax_label || "Tax"} (${Number(doc.tax_rate) || 0}%)`, money(t.tax, cur));
  }
  page.line(tx, y - 4, PAGE.W - M, y - 4, { color: INK, w: 0.9 });
  y += 6;
  row("TOTAL", money(t.total, cur), { bold: true, size: 11 });

  if (t.paid > 0) {
    row("Paid to date", money(-t.paid, cur), { color: GREEN });
    (doc.payments || []).forEach((p) => {
      const bits = [fmtDate(p.paid_on), p.method || "", p.reference || ""].filter(Boolean).join(" · ");
      page.text(tx + 10, y - 8, bits, { size: 7.5, color: MID });
      y += 10;
    });
    y += 2;
    page.line(tx, y - 4, PAGE.W - M, y - 4, { color: INK, w: 0.9 });
    y += 6;
    row("BALANCE DUE", money(t.balance, cur), { bold: true, size: 12, color: t.balance > 0 ? RED : GREEN });
  }

  y += 10;

  // ---- how to pay ----
  const payMethods: string[][] = [];
  if (doc.pay_paypal !== false && s.paypal_enabled && s.paypal_handle) {
    const pp = ["PayPal", s.paypal_handle];
    if (s.paypal_note) pp.push(s.paypal_note);
    payMethods.push(pp);
  }
  if (doc.pay_extra !== false) {
    for (const m of s.extra_methods || []) {
      if (!m || !m.label || !m.detail) continue;   // half a method is not a method
      const block = [String(m.label), String(m.detail)];
      if (m.note) block.push(String(m.note));
      payMethods.push(block);
    }
  }
  if (doc.pay_wire !== false && s.wire_enabled && (s.wire_account || s.wire_routing)) {
    const w = ["Bank transfer (ACH / wire)"];
    if (s.wire_beneficiary) w.push(`Beneficiary: ${s.wire_beneficiary}`);
    if (s.wire_bank_name) w.push(`Bank: ${s.wire_bank_name}`);
    if (s.wire_routing) w.push(`Routing (ABA): ${s.wire_routing}`);
    if (s.wire_account) w.push(`Account: ${s.wire_account}`);
    if (s.wire_swift) w.push(`SWIFT: ${s.wire_swift}`);
    if (s.wire_bank_address) w.push(s.wire_bank_address);
    if (s.wire_note) w.push(s.wire_note);
    payMethods.push(w);
  }

  if (payMethods.length && t.balance > 0) {
    const colW0 = (PAGE.W - M * 2 - 20) / payMethods.length;
    // The methods sit SIDE BY SIDE, so the block is as tall as the tallest
    // column - not the sum of them. Estimating it as a sum pushed a four-line
    // invoice onto a second page for no reason.
    const tallest = Math.max(...payMethods.map((mth) =>
      mth.slice(1).reduce((n, ln) => n + wrap(ln, 8.5, colW0).length, 0)));
    room(16 + 14 + tallest * 11 + (doc.pay_url ? 16 : 0) + 10);
    page.text(M, y, "HOW TO PAY", { size: 7, color: MID, spacing: 1.4 });
    y += 16;
    const colW = colW0;
    const startY = y;
    let maxY = y;
    payMethods.forEach((mth, i) => {
      const x = M + i * (colW + 20);
      let py = startY;
      page.text(x, py, mth[0], { size: 9.5, bold: true });
      py += 14;
      mth.slice(1).forEach((ln) => {
        wrap(ln, 8.5, colW).forEach((w2) => { page.text(x, py, w2, { size: 8.5, color: MID }); py += 11; });
      });
      maxY = Math.max(maxY, py);
    });
    y = maxY + 10;
    // No "Pay online: <url>" line. A PDF cannot make it clickable without link
    // annotations, so it printed a 36-character token for somebody to retype,
    // which is not an option at all. The email carries the link as a button,
    // and that is the only place it is actually usable.
  } else if (t.balance <= 0 && t.paid > 0) {
    page.text(M, y, "PAID IN FULL — thank you.", { size: 10, bold: true, color: GREEN, spacing: 1 });
    y += 20;
  }

  // ---- notes / terms ----
  // Both blocks measure what they are about to draw. A flat reservation pushed a
  // one-line note onto a second page that then held nothing else.
  if (doc.notes) {
    const nl = wrap(doc.notes, 9, PAGE.W - M * 2);
    room(14 + nl.length * 12 + 6);
    page.text(M, y, "NOTES", { size: 7, color: MID, spacing: 1.4 });
    y += 14;
    nl.forEach((ln) => { page.text(M, y, ln, { size: 9 }); y += 12; });
    y += 6;
  }
  if (doc.terms_text) {
    const tl = wrap(doc.terms_text, 8, PAGE.W - M * 2);
    room(tl.length * 10);
    tl.forEach((ln) => { page.text(M, y, ln, { size: 8, color: MID }); y += 10; });
  }

  // ---- footer on every page ----
  const foot = s.footer_note || `${s.biz_name || "Come With"} · ${s.biz_website || "comewith.org"}`;
  pdf.pages.forEach((p, i) => {
    p.line(M, PAGE.H - 46, PAGE.W - M, PAGE.H - 46, { color: RULE, w: 0.5 });
    p.text(M, PAGE.H - 32, foot, { size: 7.5, color: MID });
    p.text(PAGE.W - M - 120, PAGE.H - 32, `Page ${i + 1} of ${pdf.pages.length}`,
      { size: 7.5, color: MID, align: "right", width: 120 });
  });

  return pdf.build();
}

// ===========================================================================
// HTML — the same document, for the client-facing page and the email body.
// ===========================================================================
export const esc = (v: unknown) =>
  String(v ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export function invoiceHtml(doc: InvoiceDoc, opts: { standalone?: boolean; payUrl?: string | null } = {}) {
  const t = computeTotals(doc);
  const s = doc.settings || {};
  const cur = doc.currency || "USD";
  const nl2br = (v: string) => esc(v).replace(/\n/g, "<br>");

  const lineRows = (doc.lines || []).map((l) => {
    const a = lineAmounts(l);
    const note = discountNote(l);
    return `<tr>
      <td class="d"><strong>${esc(l.description)}</strong>
        ${l.detail ? `<div class="detail">${nl2br(l.detail)}</div>` : ""}
        ${note ? `<div class="disc">${esc(note)}</div>` : ""}</td>
      <td class="n">${esc(l.qty)}</td>
      <td class="n">${esc(money(Number(l.unit_price) || 0, cur))}</td>
      <td class="n">${esc(money(a.amount, cur))}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="4" class="muted">No line items.</td></tr>`;

  const totalRow = (label: string, value: string, cls = "") =>
    `<tr class="${cls}"><td colspan="2"></td><td class="lbl">${esc(label)}</td><td class="n">${esc(value)}</td></tr>`;

  let totals = totalRow("Subtotal", money(t.subtotal, cur));
  if (t.lineDiscount > 0) totals += totalRow("Line discounts", money(-t.lineDiscount, cur), "muted-row");
  if (t.invoiceDiscount > 0) {
    totals += totalRow(
      doc.discount_kind === "percent" ? `Discount (${Number(doc.discount_value)}%)` : "Discount",
      money(-t.invoiceDiscount, cur), "disc-row");
  }
  if (doc.tax_enabled) {
    totals += totalRow(`${doc.tax_label || "Tax"} (${Number(doc.tax_rate) || 0}%)`, money(t.tax, cur));
  }
  totals += totalRow("Total", money(t.total, cur), "grand");
  if (t.paid > 0) {
    totals += totalRow("Paid to date", money(-t.paid, cur), "paid-row");
    totals += totalRow("Balance due", money(t.balance, cur), "grand balance");
  }

  const payBlocks: string[] = [];
  if (doc.pay_paypal !== false && s.paypal_enabled && s.paypal_handle) {
    const handle = String(s.paypal_handle);
    const href = handle.includes("@")
      ? `https://www.paypal.com/paypalme/`
      : `https://www.paypal.com/paypalme/${encodeURIComponent(handle.replace(/^@/, ""))}${t.balance > 0 ? "/" + t.balance.toFixed(2) : ""}`;
    payBlocks.push(`<div class="pay">
      <h4>PayPal</h4>
      ${handle.includes("@")
        ? `<p>Send to <strong>${esc(handle)}</strong></p>`
        : `<p><a class="btn" href="${esc(href)}" target="_blank" rel="noopener">Pay ${esc(money(t.balance, cur))} with PayPal →</a></p>
           <p class="fine">or send to <strong>${esc(handle)}</strong></p>`}
      ${s.paypal_note ? `<p class="fine">${nl2br(String(s.paypal_note))}</p>` : ""}
    </div>`);
  }
  if (doc.pay_extra !== false) {
    for (const m of s.extra_methods || []) {
      if (!m || !m.label || !m.detail) continue;
      payBlocks.push(`<div class="pay">
        <h4>${esc(m.label)}</h4>
        <p><strong>${esc(m.detail)}</strong></p>
        ${m.note ? `<p class="fine">${nl2br(String(m.note))}</p>` : ""}
      </div>`);
    }
  }
  if (doc.pay_wire !== false && s.wire_enabled && (s.wire_account || s.wire_routing)) {
    const rows: [string, unknown][] = [
      ["Beneficiary", s.wire_beneficiary], ["Bank", s.wire_bank_name],
      ["Routing (ABA)", s.wire_routing], ["Account", s.wire_account], ["SWIFT", s.wire_swift],
    ];
    payBlocks.push(`<div class="pay">
      <h4>Bank transfer (ACH / wire)</h4>
      <table class="wire">${rows.filter(([, v]) => v).map(([k, v]) =>
        `<tr><td>${esc(k)}</td><td><strong>${esc(v)}</strong></td></tr>`).join("")}</table>
      ${s.wire_bank_address ? `<p class="fine">${nl2br(String(s.wire_bank_address))}</p>` : ""}
      ${s.wire_note ? `<p class="fine">${nl2br(String(s.wire_note))}</p>` : ""}
    </div>`);
  }

  const body = `
  <div class="inv">
    <header class="mast">
      <div>
        <div class="brand">${esc((s.biz_name || "Come With").toUpperCase())}</div>
        <div class="brand-sub">${esc([s.biz_website, s.biz_email].filter(Boolean).join("  ·  "))}</div>
      </div>
      <div class="mast-r">
        <div class="word">INVOICE</div>
        <div class="no">${esc(doc.invoice_no)}</div>
      </div>
    </header>

    <section class="meta">
      <div><h5>From</h5><p><strong>${esc(s.biz_legal_name || s.biz_name || "Come With")}</strong><br>
        ${nl2br(String(s.biz_address || ""))}${s.biz_address ? "<br>" : ""}
        ${esc(s.biz_email || "")}${s.biz_phone ? "<br>" + esc(s.biz_phone) : ""}
        ${s.tax_id ? "<br>Tax ID " + esc(s.tax_id) : ""}</p></div>
      <div><h5>Bill to</h5><p><strong>${esc(doc.bill_to_name || "—")}</strong><br>
        ${nl2br(String(doc.bill_to_address || ""))}${doc.bill_to_address ? "<br>" : ""}
        ${esc(doc.bill_to_email || "")}</p></div>
      <div class="dates">
        <h5>Issued</h5><p>${esc(fmtDate(doc.issue_date))}</p>
        <h5>Due</h5><p>${esc(doc.due_date ? fmtDate(doc.due_date) : "On receipt")}</p>
        ${doc.event_name ? `<h5>Event</h5><p>${esc(doc.event_name)}${doc.event_date ? " · " + esc(fmtDate(doc.event_date)) : ""}</p>` : ""}
      </div>
    </section>

    <div class="due ${t.state}">
      <span class="due-lbl">${t.balance > 0 ? "Amount due" : "Amount"}</span>
      <span class="due-amt">${esc(money(t.balance > 0 ? t.balance : t.total, cur))}</span>
      <span class="chip ${t.state}">${esc(STATE_LABEL[t.state] || t.state)}</span>
    </div>

    <table class="lines">
      <thead><tr><th>Description</th><th class="n">Qty</th><th class="n">Rate</th><th class="n">Amount</th></tr></thead>
      <tbody>${lineRows}</tbody>
      <tfoot>${totals}</tfoot>
    </table>

    ${payBlocks.length && t.balance > 0 ? `<section class="pays"><h3>How to pay</h3><div class="pay-grid">${payBlocks.join("")}</div></section>` : ""}
    ${t.balance <= 0 && t.paid > 0 ? `<p class="paidfull">Paid in full — thank you.</p>` : ""}
    ${doc.notes ? `<section class="notes"><h5>Notes</h5><p>${nl2br(String(doc.notes))}</p></section>` : ""}
    ${doc.terms_text ? `<p class="terms">${nl2br(String(doc.terms_text))}</p>` : ""}
    <footer class="foot">${esc(s.footer_note || `${s.biz_name || "Come With"} · ${s.biz_website || "comewith.org"}`)}</footer>
  </div>`;

  if (!opts.standalone) return body;
  return `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Invoice ${esc(doc.invoice_no)} · ${esc(s.biz_name || "Come With")}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Libre+Baskerville:wght@400;700&family=Inconsolata:wght@400;600&display=swap" rel="stylesheet">
<style>${INVOICE_CSS}</style></head><body>${body}</body></html>`;
}

export const INVOICE_CSS = `
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--ink:${INK};--bg:${CREAM};--red:${RED};--mid:${MID};--rule:${RULE};--green:${GREEN}}
body{background:var(--bg);color:var(--ink);font-family:'Libre Baskerville',Georgia,serif;line-height:1.6;padding:28px 16px 64px}
.inv{max-width:960px;margin:0 auto;background:#fff;border:1px solid var(--rule);box-shadow:0 1px 3px rgba(26,20,16,.06)}
.mast{background:var(--ink);color:var(--bg);padding:26px 34px;display:flex;justify-content:space-between;align-items:flex-start;gap:16px;border-bottom:3px solid var(--red)}
.brand{font-family:'Bebas Neue',Impact,sans-serif;font-size:2.1rem;letter-spacing:.07em;line-height:1}
.brand-sub{font-family:'Inconsolata',monospace;font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:#B9AE9E;margin-top:7px}
.mast-r{text-align:right}
.word{font-family:'Bebas Neue',Impact,sans-serif;font-size:1.9rem;letter-spacing:.14em;line-height:1}
.no{font-family:'Inconsolata',monospace;font-size:.8rem;letter-spacing:.1em;color:var(--red);margin-top:4px}
.meta{display:grid;grid-template-columns:1fr 1fr 170px;gap:22px;padding:26px 34px 6px}
.meta h5,.notes h5{font-family:'Inconsolata',monospace;font-size:.6rem;letter-spacing:.16em;text-transform:uppercase;color:var(--mid);font-weight:600;margin-bottom:5px}
.meta p{font-size:.84rem;line-height:1.5;margin-bottom:12px}
.dates p{margin-bottom:12px}
.due{display:flex;align-items:baseline;gap:12px;margin:14px 34px 0;padding:14px 18px;background:#F7F3EC;border-left:3px solid var(--ink)}
.due.overdue{border-left-color:var(--red)}
.due.paid{border-left-color:var(--green)}
.due-lbl{font-family:'Inconsolata',monospace;font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:var(--mid)}
.due-amt{font-size:1.6rem;font-weight:700}
.chip{margin-left:auto;font-family:'Inconsolata',monospace;font-size:.6rem;letter-spacing:.14em;padding:3px 10px;border:1px solid var(--mid);color:var(--mid);border-radius:2px}
.chip.paid{color:var(--green);border-color:var(--green)}
.chip.overdue{color:var(--red);border-color:var(--red)}
.chip.partial{color:#854F0B;border-color:#854F0B}
table.lines{width:100%;border-collapse:collapse;margin:22px 0 0}
table.lines th{font-family:'Inconsolata',monospace;font-size:.6rem;letter-spacing:.14em;text-transform:uppercase;color:var(--mid);font-weight:600;text-align:left;padding:0 34px 8px;border-bottom:1.5px solid var(--ink)}
table.lines th:first-child{padding-left:34px}
table.lines td{padding:11px 34px;border-bottom:1px solid var(--rule);font-size:.86rem;vertical-align:top}
table.lines .n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
table.lines th.n{text-align:right}
.detail{font-size:.76rem;color:var(--mid);margin-top:3px;line-height:1.45}
.disc{font-size:.76rem;color:var(--red);margin-top:3px}
tfoot td{border-bottom:none!important;padding-top:7px!important;padding-bottom:0!important;font-size:.84rem}
tfoot .lbl{text-align:right;color:var(--mid)}
tfoot .grand td{border-top:1.5px solid var(--ink)}
tfoot .grand .lbl,tfoot .grand .n{font-weight:700;font-size:1rem;padding-top:11px!important}
tfoot .balance .n{color:var(--red)}
tfoot .paid-row .n{color:var(--green)}
tfoot .disc-row .n{color:var(--red)}
.pays{padding:26px 34px 0;margin-top:18px;border-top:1px solid var(--rule)}
.pays h3{font-family:'Inconsolata',monospace;font-size:.62rem;letter-spacing:.16em;text-transform:uppercase;color:var(--mid);font-weight:600;margin-bottom:14px}
.pay-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:22px}
.pay h4{font-size:.92rem;margin-bottom:7px}
.pay p{font-size:.82rem;margin-bottom:6px}
.pay .fine{font-size:.75rem;color:var(--mid)}
table.wire{border-collapse:collapse;font-size:.8rem}
table.wire td{padding:2px 12px 2px 0;color:var(--mid)}
table.wire td strong{color:var(--ink);font-variant-numeric:tabular-nums}
.btn{display:inline-block;background:var(--red);color:#fff;text-decoration:none;padding:9px 17px;font-size:.82rem;font-weight:700;border-radius:2px}
.btn:hover{background:#8C2A1C}
.paidfull{padding:22px 34px;color:var(--green);font-weight:700}
.notes{padding:22px 34px 0}
.notes p{font-size:.84rem;white-space:pre-wrap}
.terms{padding:16px 34px 0;font-size:.75rem;color:var(--mid)}
.foot{padding:24px 34px 26px;margin-top:20px;border-top:1px solid var(--rule);font-family:'Inconsolata',monospace;font-size:.66rem;letter-spacing:.08em;color:var(--mid)}
.muted{color:var(--mid)}
@media print{
  body{background:#fff;padding:0}
  .inv{border:none;box-shadow:none;max-width:none}
  .noprint{display:none!important}
  @page{margin:14mm}
}
@media(max-width:640px){
  .meta{grid-template-columns:1fr}
  .mast{flex-direction:column}
  .mast-r{text-align:left}
  table.lines td,table.lines th{padding-left:18px;padding-right:18px}
}`;
