// subscribe
//
// Public endpoint called from subscribe widgets on index-v2.html,
// the DI hub page, or any future surface.
//
// Master-list architecture: there is ONE subscribers table. A given
// email is one row. Each subscribe call adds (or no-ops) a row in
// subscriber_segments. Status (subscribed/unsubscribed/etc) is global.
//
// Behavior:
//   - Email new → insert with status=pending, send double-opt-in
//     confirm email
//   - Email present + status in (pending, subscribed) → no new
//     confirm email, just add the segment
//   - Email present + status unsubscribed → flip to pending, send
//     confirm email (treats it like a fresh signup)
//   - Email present + status bounced/complained → reject (return error)
//
// Required: RESEND_API_KEY

import { createClient } from "npm:@supabase/supabase-js@2";
import { Resend } from "npm:resend@4";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

// Sender is centrally configurable via secrets; these are the fallbacks.
const FROM = Deno.env.get("FROM_EMAIL") || "Come With <berky@comewith.org>";
const REPLY_TO = Deno.env.get("REPLY_TO_EMAIL") || "berky@comewith.org";

// SITE_URL is set as a secret per project.
const SITE_URL = Deno.env.get("SITE_URL") || "http://localhost:8765";
const CONFIRM_BASE = `${SITE_URL}/confirm.html`;

function jsonError(status: number, message: string) {
  return new Response(JSON.stringify({ error: message }), { status, headers: JSON_HEADERS });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const body = await req.json().catch(() => ({}));
    const email = (body.email || "").toString().trim().toLowerCase();
    const segment = (body.segment || "main").toString().trim();
    const fullName = body.full_name ? body.full_name.toString().trim() : null;
    const source = body.source ? body.source.toString().trim() : "website";

    if (!email || !email.includes("@")) return jsonError(400, "valid email required");
    if (!segment) return jsonError(400, "segment required");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // Look up existing subscriber by lowercased email
    const { data: existing } = await admin
      .from("subscribers")
      .select("id, email, status, unsubscribe_token, confirmed_at, confirm_sent_at")
      .ilike("email", email)
      .maybeSingle();

    let subscriberId: string;
    let needsConfirm = false;
    let unsubscribeToken: string;
    let status: string;

    if (!existing) {
      // New subscriber → pending
      const { data: inserted, error: insErr } = await admin
        .from("subscribers")
        .insert({ email, full_name: fullName, status: "pending", source })
        .select("id, unsubscribe_token")
        .single();
      if (insErr || !inserted) { console.error("subscribe insert failed:", insErr?.message); return jsonError(500, "Could not subscribe right now — please try again."); }
      subscriberId = inserted.id;
      unsubscribeToken = inserted.unsubscribe_token;
      status = "pending";
      needsConfirm = true;
    } else {
      subscriberId = existing.id;
      unsubscribeToken = existing.unsubscribe_token;
      status = existing.status;

      if (status === "bounced" || status === "complained") {
        return jsonError(400, "We can't subscribe this email — please contact " + REPLY_TO + " if this is unexpected.");
      }
      if (status === "unsubscribed") {
        // Re-subscribe: flip back to pending, send a fresh confirm
        await admin
          .from("subscribers")
          .update({ status: "pending", unsubscribed_at: null })
          .eq("id", subscriberId);
        status = "pending";
        needsConfirm = true;
      }
      // If already pending or subscribed, just add the segment below.
    }

    // Add segment (idempotent via unique index)
    await admin
      .from("subscriber_segments")
      .insert({ subscriber_id: subscriberId, segment })
      .select(); // ignore result; conflict is fine

    // Rate limit: if we already sent this address a confirm in the last 10 minutes,
    // pretend success without re-sending (stops bots spamming someone's inbox).
    const lastSent = existing?.confirm_sent_at ? new Date(existing.confirm_sent_at).getTime() : 0;
    if (needsConfirm && lastSent && Date.now() - lastSent < 10 * 60 * 1000) {
      return new Response(
        JSON.stringify({ success: true, status, segment, confirm_sent: false, throttled: true }),
        { headers: JSON_HEADERS },
      );
    }

    // Send confirm email if needed
    if (needsConfirm) {
      const resendKey = Deno.env.get("RESEND_API_KEY");
      if (!resendKey) {
        return jsonError(500, "RESEND_API_KEY not set");
      }
      const resend = new Resend(resendKey);
      const confirmUrl = `${CONFIRM_BASE}?token=${unsubscribeToken}`;
      // Copy is owner-editable (email_templates key 'subscribe_confirm'); fall back
      // to the built-in text if the row is missing.
      const { data: tpl } = await admin.from("email_templates").select("subject, body").eq("key", "subscribe_confirm").maybeSingle();
      const escHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      const greeting = fullName ? "Hi " + fullName.split(" ")[0] + "," : "Hi,";
      const confirmBtn = `<a href="${confirmUrl}" style="background:#C13B2A;color:white;padding:12px 22px;text-decoration:none;font-weight:600;letter-spacing:0.04em;">Confirm subscription</a>`;
      const bodyTpl = tpl?.body || "{{greeting}}\n\nYou signed up for the Come With mailing list. One click and you're in:\n\n{{confirm_button}}\n\nIf you didn't sign up, ignore this email and you won't be subscribed.";
      const filled = escHtml(bodyTpl).replace(/\{\{\s*(\w+)\s*\}\}/g, (_, k: string) =>
        k === "confirm_button" ? confirmBtn : k === "greeting" ? escHtml(greeting) : "");
      const html = `<!doctype html><html><body style="font-family:system-ui,sans-serif;color:#1A1410;line-height:1.55;max-width:520px;margin:24px auto;padding:0 16px;">${filled.replace(/\r\n/g, "\n").replace(/\n/g, "<br>")}</body></html>`;
      const sendRes = await resend.emails.send({
        from: FROM,
        to: email,
        replyTo: REPLY_TO,
        subject: tpl?.subject || "Confirm your Come With subscription",
        html,
      });
      if (sendRes.error) {
        console.error("subscribe confirm send failed:", sendRes.error.message);
        return jsonError(500, "We couldn't send the confirmation email — please try again.");
      }
      await admin.from("subscribers").update({ confirm_sent_at: new Date().toISOString() }).eq("id", subscriberId);
    }

    return new Response(
      JSON.stringify({
        success: true,
        status,
        segment,
        confirm_sent: needsConfirm,
      }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    console.error("subscribe unexpected:", e instanceof Error ? e.message : String(e));
    return jsonError(500, "Something went wrong — please try again.");
  }
});
