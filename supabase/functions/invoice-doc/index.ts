// invoice-doc
//
// Everything that turns an invoice ROW into an invoice DOCUMENT: the PDF, the
// HTML, filing it into storage, and emailing it to the client.
//
// ONE FUNCTION, TWO AUDIENCES — deliberately.
//   admin actions  (preview | pdf | send)  need a master/sub admin JWT, or the
//                                          service-role key for internal callers
//   public actions (view | pdf_public)     need only the invoice's public_token
//
// They live together because both need `template.ts` + `pdf.ts`, and the deploy
// script ships one folder per function. Splitting them would mean two copies of
// a 700-line renderer that must never disagree about a total — the exact failure
// this whole feature is built to avoid.
//
// DEPLOY WITH JWT VERIFICATION OFF:
//   python scripts/deploy_edge_function.py invoice-doc --no-verify-jwt
// The public actions have no Supabase user; the token IS the credential. Every
// admin action re-checks the caller itself, below, so --no-verify-jwt does not
// widen anything: `requireAdmin` is the gate, not the platform.
//
// WHAT THE PUBLIC PATH MAY RETURN. Only the invoice behind the presented token,
// plus the parts of `invoice_settings` that are PRINTED ON the invoice — the
// business address and the payment instructions, wire details included. Those
// are on the document by design; that is what an invoice is. It returns no
// other invoice, no actor record, and no table the token does not name.
//
// Secrets: SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, SUPABASE_ANON_KEY,
//          RESEND_API_KEY (send only), SITE_URL, FROM_EMAIL (optional).

import { createClient } from "npm:@supabase/supabase-js@2";
import { Resend } from "npm:resend@4";
import {
  renderInvoicePdf, invoiceHtml, computeTotals, money, INVOICE_CSS,
  type InvoiceDoc, esc,
} from "./template.ts";

const HEADERS = {
  "Content-Type": "application/json",
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const err = (s: number, m: string) =>
  new Response(JSON.stringify({ error: m }), { status: s, headers: HEADERS });
const ok = (b: unknown) => new Response(JSON.stringify(b), { headers: HEADERS });

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const SITE_URL = Deno.env.get("SITE_URL") || "https://comewith.org";
const FROM = Deno.env.get("FROM_EMAIL") || "Come With <berky@comewith.org>";

const admin = () => createClient(SUPABASE_URL, SERVICE_ROLE);

/** Admin actions: a service-role bearer, or a signed-in master/sub admin. */
async function requireAdmin(req: Request): Promise<{ ok: boolean; userId?: string }> {
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  if (!bearer) return { ok: false };
  if (bearer === SERVICE_ROLE) return { ok: true };
  const userClient = createClient(SUPABASE_URL, Deno.env.get("SUPABASE_ANON_KEY")!, {
    global: { headers: { Authorization: auth } },
  });
  const { data: { user } } = await userClient.auth.getUser();
  if (!user) return { ok: false };
  const { data: prof } = await admin().from("profiles").select("role, deleted_at")
    .eq("id", user.id).single();
  // deleted_at is checked as well as role: under the 098 deactivation contract a
  // deactivated profile is no-role, and a guard that reads only `role` lets a
  // deactivated admin keep working. CLAUDE.md calls this out by name.
  const good = !!prof && !prof.deleted_at && ["master_admin", "sub_admin"].includes(prof.role);
  return { ok: good, userId: user.id };
}

const b64 = (bytes: Uint8Array) => {
  let s = "";
  const CH = 0x8000;                     // chunked: apply() blows the stack on big arrays
  for (let i = 0; i < bytes.length; i += CH) {
    s += String.fromCharCode(...bytes.subarray(i, i + CH));
  }
  return btoa(s);
};

/** Assemble the InvoiceDoc the renderers read. One loader, one shape. */
async function loadDoc(where: { id?: string; token?: string }): Promise<
  { doc: InvoiceDoc; row: Record<string, unknown> } | null
> {
  const a = admin();
  let q = a.from("invoices").select("*").is("deleted_at", null);
  q = where.token ? q.eq("public_token", where.token) : q.eq("id", where.id!);
  const { data: inv } = await q.maybeSingle();
  if (!inv) return null;

  const [{ data: lines }, { data: payments }, { data: settings }, { data: actor }, { data: event }] =
    await Promise.all([
      a.from("invoice_lines").select("*").eq("invoice_id", inv.id)
        .order("position", { ascending: true }).order("created_at", { ascending: true }),
      a.from("invoice_payments").select("*").eq("invoice_id", inv.id)
        .order("paid_on", { ascending: true }),
      a.from("invoice_settings").select("*").eq("id", true).maybeSingle(),
      inv.bill_to_actor_id
        ? a.from("actors").select("display_name, legal_name, email").eq("id", inv.bill_to_actor_id).maybeSingle()
        : Promise.resolve({ data: null }),
      inv.event_id
        ? a.from("events").select("name, event_date").eq("id", inv.event_id).maybeSingle()
        : Promise.resolve({ data: null }),
    ]);

  const doc: InvoiceDoc = {
    invoice_no: inv.invoice_no,
    status: inv.status,
    issue_date: inv.issue_date,
    due_date: inv.due_date,
    // The stored snapshot wins; the actor is only a fallback for a draft that
    // has not been sent yet. Once it has gone out, the document must not change
    // because somebody renamed a contact.
    bill_to_name: inv.bill_to_name || actor?.legal_name || actor?.display_name || null,
    bill_to_email: inv.bill_to_email || actor?.email || null,
    bill_to_address: inv.bill_to_address,
    event_name: event?.name || null,
    event_date: event?.event_date || null,
    currency: inv.currency,
    discount_kind: inv.discount_kind,
    discount_value: Number(inv.discount_value || 0),
    tax_enabled: inv.tax_enabled,
    tax_rate: Number(inv.tax_rate || 0),
    tax_label: inv.tax_label,
    notes: inv.notes,
    terms_text: inv.terms_text,
    pay_paypal: inv.pay_paypal,
    pay_wire: inv.pay_wire,
    lines: (lines || []).map((l: Record<string, unknown>) => ({
      description: l.description as string,
      detail: l.detail as string | null,
      qty: Number(l.qty),
      unit_price: Number(l.unit_price),
      discount_kind: l.discount_kind as "amount" | "percent" | null,
      discount_value: Number(l.discount_value || 0),
      taxable: l.taxable !== false,
    })),
    payments: (payments || []).map((p: Record<string, unknown>) => ({
      paid_on: p.paid_on as string,
      amount: Number(p.amount),
      method: p.method as string | null,
      reference: p.reference as string | null,
    })),
    settings: settings || {},
    pay_url: `${SITE_URL}/invoice.html?t=${inv.public_token}`,
  };
  return { doc, row: inv };
}

const fileName = (doc: InvoiceDoc) =>
  `Invoice ${doc.invoice_no}${doc.bill_to_name ? " - " + doc.bill_to_name.replace(/[\\/:*?"<>|]/g, "") : ""}.pdf`;

/** Render, upload to the private bucket, and keep the event's Files row in step. */
async function storePdf(invId: string, doc: InvoiceDoc, eventId: string | null) {
  const a = admin();
  const bytes = renderInvoicePdf(doc);
  const path = `invoice/${doc.invoice_no}.pdf`;
  const { error: upErr } = await a.storage.from("invoices")
    .upload(path, bytes, { contentType: "application/pdf", upsert: true });
  if (upErr) throw new Error("PDF upload failed: " + upErr.message);
  await a.from("invoices").update({ pdf_path: path }).eq("id", invId);

  // File it against the event, the way file-agreement does, so an invoice shows
  // up in that event's Files next to the contract. Idempotent on (bucket, path).
  if (eventId) {
    const { data: existing } = await a.from("files").select("id")
      .eq("bucket", "invoices").eq("path", path).maybeSingle();
    const row = {
      bucket: "invoices", path, filename: fileName(doc), mime: "application/pdf",
      size: bytes.length, subject_type: "event", subject_id: eventId, kind: "invoice",
    };
    if (existing) await a.from("files").update(row).eq("id", existing.id);
    else await a.from("files").insert(row);
  }
  return { bytes, path };
}

/** The email body. Short on purpose — the document is the attachment and the link. */
function emailHtml(doc: InvoiceDoc, t: ReturnType<typeof computeTotals>, note?: string | null) {
  const s = doc.settings || {};
  const due = doc.due_date
    ? new Date(doc.due_date + "T00:00:00Z").toLocaleDateString("en-US",
        { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" })
    : "on receipt";
  return `<div style="font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;color:#1A1410;max-width:560px;line-height:1.6">
  <p style="font:700 20px/1.2 Georgia,serif;margin:0 0 4px">${esc(s.biz_name || "Come With")}</p>
  <p style="color:#8A7F72;font-size:13px;margin:0 0 20px">Invoice ${esc(doc.invoice_no)}</p>
  <p>Hi${doc.bill_to_name ? " " + esc(doc.bill_to_name) : ""},</p>
  <p>${note ? esc(note) : `Please find invoice <strong>${esc(doc.invoice_no)}</strong> attached${doc.event_name ? ` for <strong>${esc(doc.event_name)}</strong>` : ""}.`}</p>
  <table style="border-collapse:collapse;margin:18px 0;font-size:15px">
    <tr><td style="padding:4px 16px 4px 0;color:#8A7F72">Amount due</td>
        <td style="padding:4px 0;font-weight:700;font-size:19px">${esc(money(t.balance, doc.currency))}</td></tr>
    <tr><td style="padding:4px 16px 4px 0;color:#8A7F72">Due</td>
        <td style="padding:4px 0">${esc(due)}</td></tr>
  </table>
  <p><a href="${esc(doc.pay_url)}" style="display:inline-block;background:#C13B2A;color:#fff;text-decoration:none;padding:11px 22px;border-radius:2px;font-weight:700">View &amp; pay online →</a></p>
  <p style="color:#8A7F72;font-size:13px;margin-top:22px">The invoice is also attached as a PDF. Reply to this email with any questions.</p>
  <p style="color:#8A7F72;font-size:12px">${esc(s.footer_note || `${s.biz_name || "Come With"} · ${s.biz_website || "comewith.org"}`)}</p>
</div>`;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: HEADERS });
  if (req.method !== "POST") return err(405, "POST only");

  const body = await req.json().catch(() => ({}));
  const action = String(body.action || "");

  // ---------------------------------------------------------------- public --
  if (action === "view" || action === "pdf_public") {
    const token = String(body.token || "");
    // A uuid or nothing. Rejecting the shape first means a malformed token never
    // reaches a query.
    if (!/^[0-9a-f-]{36}$/i.test(token)) return err(400, "bad token");
    const loaded = await loadDoc({ token });
    if (!loaded) return err(404, "not found");
    const { doc, row } = loaded;
    if (row.status === "draft" || row.status === "void") {
      // A draft has not been issued and a void one has been withdrawn. Neither
      // is a document anybody should be able to pull up from a stale link.
      return err(404, "not found");
    }
    if (action === "pdf_public") {
      const bytes = renderInvoicePdf(doc);
      return ok({ filename: fileName(doc), pdf_base64: b64(bytes) });
    }
    // First open stamps viewed_at. Never overwritten, so it records when the
    // client first saw it, not the last time anyone refreshed.
    if (!row.viewed_at) {
      await admin().from("invoices").update({ viewed_at: new Date().toISOString() }).eq("id", row.id);
    }
    // The CSS travels with the markup so invoice.html never keeps its own copy
    // of it. Two stylesheets for one document is how the emailed PDF and the
    // web page start looking like different invoices.
    return ok({
      html: invoiceHtml(doc, { standalone: false }),
      css: INVOICE_CSS,
      invoice_no: doc.invoice_no,
      biz_name: doc.settings?.biz_name || "Come With",
      totals: computeTotals(doc),
    });
  }

  // ----------------------------------------------------------------- admin --
  const gate = await requireAdmin(req);
  if (!gate.ok) return err(401, "unauthorized");

  const invoiceId = String(body.invoice_id || "");
  if (!invoiceId) return err(400, "invoice_id is required");
  const loaded = await loadDoc({ id: invoiceId });
  if (!loaded) return err(404, "Invoice not found");
  const { doc, row } = loaded;
  const t = computeTotals(doc);
  const a = admin();

  if (action === "preview") {
    return ok({ html: invoiceHtml(doc, { standalone: false }), totals: t });
  }

  if (action === "pdf") {
    const { bytes, path } = await storePdf(row.id as string, doc, (row.event_id as string) || null);
    return ok({ filename: fileName(doc), pdf_base64: b64(bytes), path, totals: t });
  }

  if (action === "send") {
    const to = String(body.to || doc.bill_to_email || "").trim();
    if (!to) return err(400, "No email address for this client — add one before sending.");
    const resendKey = Deno.env.get("RESEND_API_KEY");
    if (!resendKey) return err(500, "RESEND_API_KEY is not set on this project.");

    const { bytes, path } = await storePdf(row.id as string, doc, (row.event_id as string) || null);

    const resend = new Resend(resendKey);
    const subject = body.subject
      ? String(body.subject)
      : `Invoice ${doc.invoice_no} from ${doc.settings?.biz_name || "Come With"}` +
        (t.balance > 0 ? ` — ${money(t.balance, doc.currency)} due` : "");
    const sent = await resend.emails.send({
      from: FROM,
      to: [to],
      cc: body.cc ? [String(body.cc)] : undefined,
      reply_to: doc.settings?.biz_email || undefined,
      subject,
      html: emailHtml(doc, t, body.note),
      attachments: [{ filename: fileName(doc), content: b64(bytes) }],
    });
    if (sent.error) return err(502, "Resend refused it: " + sent.error.message);

    // Only now does the invoice count as issued.
    const patch: Record<string, unknown> = { pdf_path: path };
    if (row.status === "draft") { patch.status = "sent"; patch.sent_at = new Date().toISOString(); }
    else if (!row.sent_at) patch.sent_at = new Date().toISOString();
    // Re-sending a reminder must not reset the clock on an invoice already out.
    await a.from("invoices").update(patch).eq("id", row.id);

    // The income rows this invoice bills move accrued -> invoiced. This is the
    // state 161 defined and nothing could reach until now. Rows already
    // 'received' are left alone: money that has landed cannot un-land because a
    // document was re-sent.
    const { data: ls } = await a.from("invoice_lines").select("income_id")
      .eq("invoice_id", row.id).not("income_id", "is", null);
    const incomeIds = (ls || []).map((l: { income_id: string }) => l.income_id);
    if (incomeIds.length) {
      await a.from("income").update({ status: "invoiced" })
        .in("id", incomeIds).eq("status", "accrued");
    }
    return ok({ sent: true, to, resend_id: sent.data?.id, path, invoiced: incomeIds.length });
  }

  return err(400, "unknown action");
});
