// inquiry-notify
//
// Public endpoint called by index-v2.html immediately after the public
// inquiry form INSERTs to public.inquiries. Looks up the most recent
// matching inquiry by email (so the email contents come from the DB,
// not arbitrary frontend input) and emails every master_admin with the
// details. Fire-and-forget on the frontend side — if this fails, the
// inquiry still saved and is visible in the dashboard.
//
// Required Edge Function secret: RESEND_API_KEY

import { createClient } from "npm:@supabase/supabase-js@2";
import { Resend } from "npm:resend@4";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

const FROM = Deno.env.get("FROM_EMAIL") || "Come With <berky@comewith.org>";
const REPLY_TO = Deno.env.get("REPLY_TO_EMAIL") || "berky@comewith.org";

function jsonError(status: number, message: string) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: JSON_HEADERS,
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const body = await req.json().catch(() => ({}));
    const email = (body.email || "").toString().trim().toLowerCase();
    if (!email) return jsonError(400, "email required");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // Find the most recent inquiry from this email in the last 5 minutes.
    // The 5-minute window keeps an attacker from triggering notifications
    // for ANY past inquiry by replaying the email value.
    const fiveMinAgo = new Date(Date.now() - 5 * 60 * 1000).toISOString();
    const { data: inquiry, error: lookupErr } = await admin
      .from("inquiries")
      .select("id, full_name, email, phone, event_type, event_date, venue, services_selected, message, source, created_at")
      .ilike("email", email)
      .gte("created_at", fiveMinAgo)
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (lookupErr) { console.error("inquiry-notify lookup failed:", lookupErr.message); return jsonError(500, "Could not process the notification."); }
    if (!inquiry) {
      return jsonError(404, "No matching recent inquiry found");
    }

    // Rate limit: if this email has filed more than 3 inquiries in the past hour,
    // save them (already inserted) but skip the admin notification — stops a bot
    // from turning the public form into an email cannon at the admins.
    const hourAgo = new Date(Date.now() - 60 * 60 * 1000).toISOString();
    const { count: recentCount } = await admin
      .from("inquiries")
      .select("id", { count: "exact", head: true })
      .ilike("email", email)
      .gte("created_at", hourAgo);
    if ((recentCount || 0) > 3) {
      return new Response(
        JSON.stringify({ success: true, inquiry_id: inquiry.id, notified: 0, throttled: true }),
        { headers: JSON_HEADERS },
      );
    }

    // Look up all master_admin recipients
    const { data: admins } = await admin
      .from("profiles")
      .select("email")
      .eq("role", "master_admin");
    const recipients = (admins || []).map((a) => a.email).filter(Boolean) as string[];

    if (recipients.length === 0) {
      return new Response(
        JSON.stringify({ success: true, inquiry_id: inquiry.id, notified: 0, note: "No master_admin to notify" }),
        { headers: JSON_HEADERS },
      );
    }

    const resendKey = Deno.env.get("RESEND_API_KEY");
    if (!resendKey) {
      return jsonError(500, "RESEND_API_KEY not set");
    }

    // Build the email body
    const services = Array.isArray(inquiry.services_selected) && inquiry.services_selected.length
      ? (inquiry.services_selected as unknown[]).map((s) => typeof s === "string" ? s : JSON.stringify(s)).join(", ")
      : "—";

    const dataRow = (k: string, v: unknown) =>
      v ? `<tr><td style="padding:6px 14px 6px 0;color:#8A7F72;font-size:0.75rem;letter-spacing:0.12em;text-transform:uppercase;">${k}</td><td style="padding:6px 0;">${String(v)}</td></tr>` : "";

    const html = `<!doctype html><html><body style="font-family:system-ui,sans-serif;color:#1A1410;line-height:1.55;max-width:600px;margin:24px auto;padding:0 16px;">
      <h1 style="font-size:1.3rem;letter-spacing:0.02em;border-bottom:2px solid #C13B2A;padding-bottom:8px;">New inquiry: ${inquiry.full_name || "(no name)"}</h1>
      <table style="border-collapse:collapse;width:100%;margin-top:16px;">
        ${dataRow("Email", inquiry.email)}
        ${dataRow("Phone", inquiry.phone)}
        ${dataRow("Event type", inquiry.event_type)}
        ${dataRow("Event date", inquiry.event_date)}
        ${dataRow("Venue", inquiry.venue)}
        ${dataRow("Services", services)}
        ${dataRow("Source", inquiry.source)}
        ${dataRow("Submitted", inquiry.created_at)}
      </table>
      ${inquiry.message ? `<div style="margin-top:18px;padding:14px 16px;background:#E8E2D9;border-left:3px solid #1A1410;">
        <div style="font-size:0.7rem;letter-spacing:0.14em;text-transform:uppercase;color:#8A7F72;margin-bottom:6px;">Message</div>
        ${String(inquiry.message).replace(/</g, "&lt;").replace(/\n/g, "<br>")}
      </div>` : ""}
      <p style="margin-top:24px;font-size:0.85rem;">Reply to ${inquiry.email} or open the dashboard to update status.</p>
    </body></html>`;

    const resend = new Resend(resendKey);
    const sendRes = await resend.emails.send({
      from: FROM,
      to: recipients,
      replyTo: inquiry.email || REPLY_TO, // hitting Reply goes to the inquirer
      subject: `New inquiry — ${inquiry.full_name || inquiry.email}`,
      html,
    });

    if (sendRes.error) {
      console.error("inquiry-notify send failed:", sendRes.error.message);
      return jsonError(500, "Could not send the notification.");
    }

    return new Response(
      JSON.stringify({
        success: true,
        inquiry_id: inquiry.id,
        notified: recipients.length,
        resend_id: sendRes.data?.id,
      }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    console.error("inquiry-notify unexpected:", e instanceof Error ? e.message : String(e));
    return jsonError(500, "Something went wrong.");
  }
});
