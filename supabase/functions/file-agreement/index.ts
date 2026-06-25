// file-agreement
//
// Generates a self-contained HTML snapshot of a (signed) client agreement and
// files it into the linked event's Files tab as a `files` row
// (subject_type='event', subject_id=event_id, kind='contract'), so a signed
// client agreement shows up automatically under that event's Contract bucket.
//
// Idempotent: a deterministic storage path (event-agreement/<id>.html) is
// upserted, and the matching files row is reused if present.
//
// Auth: the service-role key (used by mark-signed) OR a valid master/sub admin
// JWT (used by the dashboard). Returns { filed:false } when no event is linked.
// Secrets: the standard SUPABASE_* set (no extra secrets needed).

import { createClient } from "npm:@supabase/supabase-js@2";

const HEADERS = {
  "Content-Type": "application/json",
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: HEADERS });
const esc = (s: unknown) =>
  String(s ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
const money = (n: unknown) =>
  n == null || n === "" ? "—" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: HEADERS });
  if (req.method !== "POST") return err(405, "POST only");

  const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
  const SERVICE_ROLE = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

  // --- auth: service-role bearer OR an admin session ---
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  let authed = bearer === SERVICE_ROLE;
  if (!authed && bearer) {
    const userClient = createClient(SUPABASE_URL, Deno.env.get("SUPABASE_ANON_KEY")!, {
      global: { headers: { Authorization: auth } },
    });
    const { data: { user } } = await userClient.auth.getUser();
    if (user) {
      const a = createClient(SUPABASE_URL, SERVICE_ROLE);
      const { data: prof } = await a.from("profiles").select("role").eq("id", user.id).single();
      authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
    }
  }
  if (!authed) return err(401, "unauthorized");

  const body = await req.json().catch(() => ({}));
  const agreementId = body.agreement_id;
  if (!agreementId) return err(400, "agreement_id is required");

  const admin = createClient(SUPABASE_URL, SERVICE_ROLE);

  const { data: ag, error: agErr } = await admin
    .from("agreements")
    .select(
      "id, agreement_type, status, event_id, event_date, venue_name, venue_address, subtotal, deposit_amount, total_amount, payment_method, notes, client_signature_url, client_signed_at, admin_signature_url, admin_signed_at, actor:actors(display_name, email), event:events(name, event_date)",
    )
    .eq("id", agreementId)
    .single();
  if (agErr || !ag) return err(404, "Agreement not found");
  if (!ag.event_id) return new Response(JSON.stringify({ filed: false, reason: "no event linked" }), { headers: HEADERS });

  // --- build a printable HTML snapshot of the agreement ---
  const client = ag.actor?.display_name || "Client";
  const sig = ag.client_signature_url
    ? `Signed by <strong>${esc(ag.client_signature_url)}</strong> on ${esc((ag.client_signed_at || "").slice(0, 10))}`
    : "Not yet signed";
  const adminSig = ag.admin_signature_url
    ? `<p>Come With countersigned by <strong>${esc(ag.admin_signature_url)}</strong> on ${esc((ag.admin_signed_at || "").slice(0, 10))}</p>`
    : "";
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>Agreement — ${esc(client)}</title></head>
<body style="font-family:system-ui,Segoe UI,sans-serif;color:#1A1410;line-height:1.55;max-width:680px;margin:32px auto;padding:0 20px;">
  <h1 style="font-size:1.4rem;color:#3B6D11;margin-bottom:4px;">Come With — ${esc(ag.agreement_type)} agreement</h1>
  <p style="color:#8A7F72;margin-top:0;">Event: ${esc(ag.event?.name || "—")}${ag.event?.event_date ? " · " + esc(String(ag.event.event_date).slice(0, 10)) : ""} · Status: ${esc(ag.status)}</p>
  <table style="border-collapse:collapse;width:100%;margin:16px 0;font-size:0.95rem;">
    <tbody>
      <tr><td style="padding:6px 8px;color:#8A7F72;width:160px;">Client</td><td style="padding:6px 8px;">${esc(client)}${ag.actor?.email ? " · " + esc(ag.actor.email) : ""}</td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;">Event date</td><td style="padding:6px 8px;">${esc(String(ag.event_date || "—").slice(0, 10))}</td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;">Venue</td><td style="padding:6px 8px;">${esc(ag.venue_name || "—")}${ag.venue_address ? "<br>" + esc(ag.venue_address) : ""}</td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;">Subtotal</td><td style="padding:6px 8px;">${money(ag.subtotal)}</td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;">Deposit</td><td style="padding:6px 8px;">${money(ag.deposit_amount)}</td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;"><strong>Total</strong></td><td style="padding:6px 8px;"><strong>${money(ag.total_amount)}</strong></td></tr>
      <tr><td style="padding:6px 8px;color:#8A7F72;">Payment</td><td style="padding:6px 8px;">${esc(ag.payment_method || "—")}</td></tr>
    </tbody>
  </table>
  ${ag.notes ? `<p style="white-space:pre-wrap;background:#F7F4EF;padding:12px;border-radius:8px;">${esc(ag.notes)}</p>` : ""}
  <hr style="border:none;border-top:1px solid #E5DED3;margin:20px 0;">
  <p>${sig}</p>
  ${adminSig}
  <p style="font-size:0.8rem;color:#8A7F72;">Filed automatically from the Come With dashboard. This is a snapshot of the agreement record.</p>
</body></html>`;

  const bytes = new TextEncoder().encode(html);
  const path = `event-agreement/${ag.id}.html`;
  const { error: upErr } = await admin.storage.from("agreements").upload(path, bytes, { contentType: "text/html", upsert: true });
  if (upErr) return err(500, "Snapshot upload failed: " + upErr.message);

  const safeDate = String(ag.event_date || ag.event?.event_date || "").slice(0, 10);
  const filename = `Agreement - ${client}${safeDate ? " - " + safeDate : ""}.html`;

  // Reuse the existing files row for this snapshot if present (idempotent).
  const { data: existing } = await admin
    .from("files")
    .select("id")
    .eq("bucket", "agreements").eq("path", path).maybeSingle();

  let fileId: string;
  if (existing) {
    await admin.from("files").update({
      filename, mime: "text/html", size: bytes.length, subject_type: "event", subject_id: ag.event_id, kind: "contract",
    }).eq("id", existing.id);
    fileId = existing.id;
  } else {
    const { data: frow, error: insErr } = await admin.from("files").insert({
      bucket: "agreements", path, filename, mime: "text/html", size: bytes.length,
      subject_type: "event", subject_id: ag.event_id, kind: "contract",
    }).select("id").single();
    if (insErr) return err(500, "Filing record failed: " + insErr.message);
    fileId = frow.id;
  }

  return new Response(JSON.stringify({ filed: true, file_id: fileId, event_id: ag.event_id, path }), { headers: HEADERS });
});
