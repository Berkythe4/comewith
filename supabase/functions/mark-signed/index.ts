// mark-signed
//
// Public endpoint called by sign.html on form submit. Validates the
// token, marks the agreement signed, marks the link used, and emails
// every master_admin in the system a notification.
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

const FROM = "Come With <berky@comewith.org>";
const REPLY_TO = "berky@comewith.org";

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
    const token = body.token;
    const signature_name = (body.signature_name || "").toString().trim();
    if (!token || typeof token !== "string") return jsonError(400, "token is required");
    if (signature_name.length < 2) return jsonError(400, "signature_name must be at least 2 characters");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // ---- Validate token and lookup agreement ----
    const { data: link, error: linkErr } = await admin
      .from("agreement_links")
      .select("agreement_id, expires_at, used_at")
      .eq("token", token)
      .single();

    if (linkErr || !link) return jsonError(404, "Invalid signing link");
    if (new Date(link.expires_at) < new Date()) {
      return jsonError(410, "This signing link has expired.");
    }
    if (link.used_at) {
      return jsonError(409, "This agreement has already been signed.");
    }

    const { data: agreement, error: agErr } = await admin
      .from("agreements")
      .select(
        "id, agreement_type, event_date, venue_name, total_amount, status, event_id, actor:actors(display_name, email)",
      )
      .eq("id", link.agreement_id)
      .single();

    if (agErr || !agreement) return jsonError(404, "Agreement not found");

    // ---- Record the signature ----
    const now = new Date().toISOString();
    const { error: updErr } = await admin
      .from("agreements")
      .update({
        client_signed_at: now,
        client_signature_url: signature_name, // typed-name signature for now
        status: "signed",
      })
      .eq("id", link.agreement_id);

    if (updErr) return jsonError(500, "Could not mark signed: " + updErr.message);

    await admin
      .from("agreement_links")
      .update({ used_at: now })
      .eq("token", token);

    // ---- Auto-file the signed snapshot into the linked event's Files (non-fatal) ----
    let filed = false;
    if (agreement.event_id) {
      try {
        const fr = await fetch(`${Deno.env.get("SUPABASE_URL")}/functions/v1/file-agreement`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")}`,
          },
          body: JSON.stringify({ agreement_id: link.agreement_id }),
        });
        filed = fr.ok && !!(await fr.json().catch(() => ({})))?.filed;
      } catch (_) { /* signing must succeed even if filing fails */ }
    }

    // ---- Notify all master_admins ----
    const { data: admins } = await admin
      .from("profiles")
      .select("email, full_name")
      .eq("role", "master_admin");

    const recipients = (admins || [])
      .map((a) => a.email)
      .filter(Boolean) as string[];

    let notifyResult: { ok: boolean; detail?: string } = { ok: true };
    if (recipients.length > 0) {
      const resendKey = Deno.env.get("RESEND_API_KEY");
      if (resendKey) {
        const resend = new Resend(resendKey);
        const clientName = agreement.actor?.display_name || "the client";
        const eventLine = agreement.event_date ? `<p>Event: ${agreement.event_date}</p>` : "";
        const venueLine = agreement.venue_name ? `<p>Venue: ${agreement.venue_name}</p>` : "";
        const totalLine = agreement.total_amount
          ? `<p>Total: $${Number(agreement.total_amount).toFixed(2)}</p>`
          : "";
        const html = `<!doctype html><html><body style="font-family:system-ui,sans-serif;color:#1A1410;line-height:1.55;max-width:560px;margin:24px auto;padding:0 16px;">
          <h1 style="font-size:1.3rem;color:#3B6D11;">Agreement signed ✓</h1>
          <p><strong>${clientName}</strong> signed the ${agreement.agreement_type} agreement as <strong>${signature_name}</strong>.</p>
          ${eventLine}${venueLine}${totalLine}
          <p style="font-size:0.85rem;color:#8A7F72;">Signed at ${now}</p>
        </body></html>`;

        const sendRes = await resend.emails.send({
          from: FROM,
          to: recipients,
          replyTo: REPLY_TO,
          subject: `Agreement signed by ${signature_name}`,
          html,
        });

        if (sendRes.error) {
          notifyResult = { ok: false, detail: sendRes.error.message };
        }
      } else {
        notifyResult = { ok: false, detail: "RESEND_API_KEY not set" };
      }
    }

    return new Response(
      JSON.stringify({
        success: true,
        signed_at: now,
        signed_as: signature_name,
        filed_to_event: filed,
        admin_notify: notifyResult,
      }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
