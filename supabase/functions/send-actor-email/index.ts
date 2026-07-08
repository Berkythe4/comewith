// send-actor-email
//
// Admin-only. Emails one or more actors / venues (or raw addresses) via Resend
// and logs each as a Conversation thread + outbound message. Supports replying
// into an existing thread. The subject is tagged with WHERE it was sent from,
// and the body gets a deep link back to that page in the dashboard.
//
// Bounces / deliveries are tracked by resend-webhook (correlates on resend_id).
// Auth: caller's JWT must be master_admin or sub_admin.
// Secrets: RESEND_API_KEY (+ the standard SUPABASE_* set).

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const FROM = Deno.env.get("FROM_EMAIL") || "Come With <berky@comewith.org>";
const REPLY_TO = Deno.env.get("REPLY_TO_EMAIL") || "berky@comewith.org";
const SITE = "https://comewith.org";
const esc = (s: unknown) =>
  String(s ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

function gotoLink(kind: string | null, id: string | null) {
  if (!kind || !id) return `${SITE}/dashboard.html`;
  return `${SITE}/dashboard.html?goto=${encodeURIComponent(kind)}&id=${encodeURIComponent(id)}`;
}

async function sendResend(apiKey: string, to: string, subject: string, html: string, refId: string) {
  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { "Authorization": `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({ from: FROM, to, reply_to: REPLY_TO, subject, html, headers: { "X-Entity-Ref-ID": refId } }),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok) return { ok: false, id: null, error: j.message || `Resend ${res.status}` };
  return { ok: true, id: j.id as string, error: null };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const URL = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const apiKey = Deno.env.get("RESEND_API_KEY");
  if (!apiKey) return err(500, "RESEND_API_KEY not set");

  // --- auth: caller must be master_admin or sub_admin ---
  const authHeader = req.headers.get("Authorization") || "";
  if (!authHeader) return err(401, "Missing Authorization");
  const caller = createClient(URL, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: authHeader } } });
  const { data: { user } } = await caller.auth.getUser();
  if (!user) return err(401, "Invalid session");
  const admin = createClient(URL, SRK);
  const { data: prof } = await admin.from("profiles").select("role").eq("id", user.id).single();
  if (!prof || !["master_admin", "sub_admin"].includes(prof.role)) return err(403, "Admins only");

  const b = await req.json().catch(() => ({}));
  const subject = (b.subject || "").toString().trim();
  const bodyHtml = (b.body || "").toString();
  if (!subject) return err(400, "subject required");

  // ---- REPLY into an existing thread ----
  if (b.conversation_id) {
    const { data: conv } = await admin.from("conversations").select("*").eq("id", b.conversation_id).single();
    if (!conv) return err(404, "Conversation not found");
    const link = gotoLink(conv.source_kind, conv.source_id);
    const html = `${bodyHtml}<hr style="border:none;border-top:1px solid #eee;margin:18px 0;"><p style="font-size:.85rem;"><a href="${link}">Open in Come With dashboard →</a></p>`;
    const r = await sendResend(apiKey, conv.recipient_email, subject, html, b.conversation_id);
    const { data: msg } = await admin.from("conversation_messages").insert({
      conversation_id: conv.id, direction: "outbound", from_email: REPLY_TO, to_email: conv.recipient_email,
      body: bodyHtml, subject_line: subject, resend_id: r.id, status: r.ok ? "sent" : "failed",
      created_by: user.id, meta: r.ok ? {} : { error: r.error },
    } as Record<string, unknown>).select("id").single();
    await admin.from("conversations").update({ last_message_at: new Date().toISOString() }).eq("id", conv.id);
    return new Response(JSON.stringify({ ok: r.ok, conversation_id: conv.id, message_id: msg?.id, error: r.error }), { headers: JH });
  }

  // ---- NEW: one thread per recipient ----
  const recips = Array.isArray(b.recipients) ? b.recipients : [];
  if (!recips.length) return err(400, "recipients required");
  const source = (b.source || "").toString();        // human label
  const sourceKind = b.source_kind || null;           // actor | venue | event_people
  const sourceId = b.source_id || null;
  const eventId = b.event_id || null;
  const visibility = b.visibility === "restricted" ? "restricted" : "team";
  const aclIds: string[] = Array.isArray(b.acl_user_ids) ? b.acl_user_ids : [];
  const taggedSubject = source ? `[${source}] ${subject}` : subject;

  const results: unknown[] = [];
  for (const r of recips) {
    let email = (r.email || "").toString().trim();
    let actorId = r.actor_id || null;
    let name = r.name || "";
    if (!email && actorId) {
      const { data: a } = await admin.from("actors").select("email, display_name").eq("id", actorId).single();
      email = a?.email || ""; name = name || a?.display_name || "";
    }
    if (!email && r.venue_id) {
      const { data: v } = await admin.from("venues").select("contact_email, actor_id, name").eq("id", r.venue_id).single();
      email = v?.contact_email || ""; name = name || v?.name || ""; actorId = actorId || v?.actor_id || null;
    }
    if (!email) { results.push({ recipient: r, ok: false, error: "no email on record" }); continue; }

    const { data: conv, error: convErr } = await admin.from("conversations").insert({
      subject, actor_id: actorId, recipient_email: email, source, source_kind: sourceKind, source_id: sourceId,
      event_id: eventId, created_by: user.id, visibility,
    }).select("id").single();
    if (convErr || !conv) { results.push({ recipient: r, ok: false, error: convErr?.message || "thread create failed" }); continue; }

    if (visibility === "restricted" && aclIds.length) {
      await admin.from("conversation_acl").insert(aclIds.map((u) => ({ conversation_id: conv.id, user_id: u })));
    }
    const link = gotoLink(sourceKind, sourceId);
    const html = `${bodyHtml}<hr style="border:none;border-top:1px solid #eee;margin:18px 0;"><p style="font-size:.85rem;color:#777;">Sent from <strong>${esc(source || "Come With")}</strong>. <a href="${link}">Open in Come With dashboard →</a></p>`;
    const sent = await sendResend(apiKey, email, taggedSubject, html, conv.id);
    await admin.from("conversation_messages").insert({
      conversation_id: conv.id, direction: "outbound", from_email: REPLY_TO, to_email: email,
      body: bodyHtml, subject_line: taggedSubject, resend_id: sent.id, status: sent.ok ? "sent" : "failed",
      created_by: user.id, meta: sent.ok ? {} : { error: sent.error },
    } as Record<string, unknown>);
    results.push({ recipient: r, email, ok: sent.ok, conversation_id: conv.id, resend_id: sent.id, error: sent.error });
  }
  const okCount = results.filter((x: any) => x.ok).length;
  return new Response(JSON.stringify({ ok: okCount > 0, sent: okCount, total: results.length, results }), { headers: JH });
});
