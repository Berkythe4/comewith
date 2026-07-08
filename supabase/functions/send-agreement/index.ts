// send-agreement
//
// Called by the dashboard with { agreement_id } in the body.
// Creates a fresh agreement_links token, emails the customer the
// signing URL via Resend, and updates the agreement status to 'sent'.
//
// Auth: requires the caller to be an authenticated admin
// (master_admin or sub_admin). Verified by checking the JWT against
// public.profiles.role.
//
// Required Edge Function secret: RESEND_API_KEY
//   set via: supabase secrets set RESEND_API_KEY=re_xxx

import { createClient } from "npm:@supabase/supabase-js@2";
import { Resend } from "npm:resend@4";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

// SITE_URL is set as a secret per project: staging defaults to localhost,
// prod sets it to https://comewith.org via `supabase secrets set SITE_URL=...`
const SITE_URL = Deno.env.get("SITE_URL") || "http://localhost:8765";
const SIGN_BASE_URL = `${SITE_URL}/sign.html`;

const FROM = "Berky <berky@comewith.org>";
const REPLY_TO = "berky@comewith.org";

function jsonError(status: number, message: string) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: JSON_HEADERS,
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }
  if (req.method !== "POST") {
    return jsonError(405, "Method not allowed");
  }

  try {
    // ---- Verify the caller is an admin via their JWT ----
    const authHeader = req.headers.get("Authorization");
    if (!authHeader) return jsonError(401, "Missing Authorization header");

    const userClient = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: authHeader } } },
    );

    const { data: { user }, error: userErr } = await userClient.auth.getUser();
    if (userErr || !user) return jsonError(401, "Invalid session");

    // Service-role client bypasses RLS for everything below.
    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { data: profile } = await admin
      .from("profiles")
      .select("role")
      .eq("id", user.id)
      .single();

    if (!profile || !["master_admin", "sub_admin"].includes(profile.role)) {
      return jsonError(403, "Admin only");
    }

    // ---- Parse and validate the payload ----
    const body = await req.json().catch(() => ({}));
    const agreement_id = body.agreement_id;
    if (!agreement_id) return jsonError(400, "agreement_id is required");

    // ---- Fetch the agreement + its client ----
    const { data: agreement, error: agErr } = await admin
      .from("agreements")
      .select(
        "id, agreement_type, status, event_date, total_amount, venue_name, client:clients(full_name, email)",
      )
      .eq("id", agreement_id)
      .single();

    if (agErr || !agreement) return jsonError(404, "Agreement not found");
    if (!agreement.client?.email) {
      return jsonError(400, "Agreement's client has no email on file");
    }

    // ---- Create a fresh agreement_link token ----
    const { data: link, error: linkErr } = await admin
      .from("agreement_links")
      .insert({ agreement_id })
      .select("token, expires_at")
      .single();

    if (linkErr || !link) {
      return jsonError(500, "Could not create signing link: " + linkErr?.message);
    }

    // ---- Send the email via Resend ----
    const resendKey = Deno.env.get("RESEND_API_KEY");
    if (!resendKey) {
      return jsonError(
        500,
        "RESEND_API_KEY not set. Run: supabase secrets set RESEND_API_KEY=re_xxx",
      );
    }

    const resend = new Resend(resendKey);
    const signUrl = `${SIGN_BASE_URL}?token=${link.token}`;
    const clientName = agreement.client.full_name || "there";
    const totalLine = agreement.total_amount
      ? `<p>Total: $${Number(agreement.total_amount).toFixed(2)}</p>`
      : "";
    const eventLine = agreement.event_date
      ? `<p>Event date: ${agreement.event_date}</p>`
      : "";
    const venueLine = agreement.venue_name
      ? `<p>Venue: ${agreement.venue_name}</p>`
      : "";

    const html = `<!doctype html><html><body style="font-family:system-ui,sans-serif;color:#1A1410;line-height:1.55;max-width:560px;margin:24px auto;padding:0 16px;">
      <h1 style="font-size:1.4rem;font-weight:600;letter-spacing:0.02em;color:#1A1410;">Your ${agreement.agreement_type} agreement</h1>
      <p>Hi ${clientName},</p>
      <p>Please review and sign your agreement at the link below:</p>
      ${eventLine}${venueLine}${totalLine}
      <p style="margin:28px 0;">
        <a href="${signUrl}" style="background:#C13B2A;color:white;padding:12px 22px;text-decoration:none;font-weight:600;letter-spacing:0.04em;">Review &amp; sign</a>
      </p>
      <p style="font-size:0.9rem;color:#8A7F72;">Or copy this link into your browser:<br><span style="word-break:break-all;">${signUrl}</span></p>
      <p style="font-size:0.9rem;color:#8A7F72;">This link expires in 30 days.</p>
      <p>— Berky<br>Come With</p>
    </body></html>`;

    const emailRes = await resend.emails.send({
      from: FROM,
      to: agreement.client.email,
      replyTo: REPLY_TO,
      subject: `Your agreement with Come With · review &amp; sign`,
      html,
    });

    if (emailRes.error) {
      return jsonError(500, "Email send failed: " + emailRes.error.message);
    }

    // ---- Bump the agreement status to 'sent' ----
    // The email is already out at this point, so a failed status write must NOT
    // read as a failed send — surface it as a warning so the dashboard can say
    // "sent, but flip the status manually".
    const { error: statusErr } = await admin
      .from("agreements")
      .update({ status: "sent" })
      .eq("id", agreement_id);

    return new Response(
      JSON.stringify({
        success: true,
        sent_to: agreement.client.email,
        token: link.token,
        sign_url: signUrl,
        resend_id: emailRes.data?.id,
        warning: statusErr ? "Email sent, but the agreement status could not be updated to 'sent' — set it manually." : undefined,
      }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
