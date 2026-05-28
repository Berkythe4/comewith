// send-campaign
//
// Admin-only. Takes {campaign_id}, loads the campaign + filters
// subscribers by segment + status=subscribed, sends individually
// via Resend with per-recipient unsubscribe links. Inserts a
// mailing_events row per send (resend_event_id only filled by
// the webhook later).
//
// Updates the campaign: status=sending → sent (or failed),
// recipient_count, sent_at.
//
// Required secret: RESEND_API_KEY

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

// SITE_URL is set as a secret per project.
const SITE_URL = Deno.env.get("SITE_URL") || "http://localhost:8765";
const UNSUB_BASE = `${SITE_URL}/unsubscribe.html`;

function jsonError(s: number, m: string) {
  return new Response(JSON.stringify({ error: m }), { status: s, headers: JSON_HEADERS });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    // Verify caller is admin
    const auth = req.headers.get("Authorization");
    if (!auth) return jsonError(401, "auth required");
    const userClient = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: auth } } },
    );
    const { data: { user } } = await userClient.auth.getUser();
    if (!user) return jsonError(401, "invalid session");

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
      return jsonError(403, "admin only");
    }

    const body = await req.json().catch(() => ({}));
    const campaign_id = (body.campaign_id || "").toString().trim();
    if (!campaign_id) return jsonError(400, "campaign_id required");

    // Load campaign
    const { data: campaign, error: cErr } = await admin
      .from("mailing_campaigns")
      .select("id, name, subject, preview_text, body_html, body_text, segment_filter, status")
      .eq("id", campaign_id)
      .single();
    if (cErr || !campaign) return jsonError(404, "campaign not found");
    if (campaign.status === "sent") return jsonError(409, "campaign already sent");
    if (campaign.status === "sending") return jsonError(409, "campaign send in progress");

    // Mark sending
    await admin.from("mailing_campaigns").update({ status: "sending" }).eq("id", campaign_id);

    // Pick recipients
    // segment_filter is a free-text string. Empty = send to all subscribed.
    // If set, only subscribers in that segment.
    let recipientQuery = admin
      .from("subscribers")
      .select("id, email, full_name, unsubscribe_token")
      .eq("status", "subscribed");

    if (campaign.segment_filter) {
      // Filter to subscribers with this segment
      const { data: segRows } = await admin
        .from("subscriber_segments")
        .select("subscriber_id")
        .eq("segment", campaign.segment_filter);
      const ids = (segRows || []).map((r) => r.subscriber_id);
      if (ids.length === 0) {
        await admin.from("mailing_campaigns").update({
          status: "sent", recipient_count: 0, sent_at: new Date().toISOString(),
        }).eq("id", campaign_id);
        return new Response(JSON.stringify({ success: true, sent: 0, note: "no recipients in segment" }), { headers: JSON_HEADERS });
      }
      recipientQuery = recipientQuery.in("id", ids);
    }

    const { data: recipients, error: rErr } = await recipientQuery;
    if (rErr) {
      await admin.from("mailing_campaigns").update({ status: "failed" }).eq("id", campaign_id);
      return jsonError(500, "couldn't load recipients: " + rErr.message);
    }

    // Send
    const resendKey = Deno.env.get("RESEND_API_KEY");
    if (!resendKey) {
      await admin.from("mailing_campaigns").update({ status: "failed" }).eq("id", campaign_id);
      return jsonError(500, "RESEND_API_KEY not set");
    }
    const resend = new Resend(resendKey);

    let sent = 0;
    let failed = 0;
    for (const r of recipients || []) {
      const unsubUrl = `${UNSUB_BASE}?token=${r.unsubscribe_token}`;
      const footer = `<hr style="border:none;border-top:1px solid #ddd;margin:32px 0 12px;"><p style="font-size:0.75rem;color:#8A7F72;text-align:center;">You're getting this because you subscribed to Come With updates. <a href="${unsubUrl}" style="color:#8A7F72;">Unsubscribe</a> anytime.</p>`;
      const html = (campaign.body_html || campaign.body_text || "") + footer;

      const sendRes = await resend.emails.send({
        from: FROM,
        to: r.email,
        replyTo: REPLY_TO,
        subject: campaign.subject,
        html,
        headers: { "List-Unsubscribe": `<${unsubUrl}>` },
      });

      if (sendRes.error) {
        failed++;
        await admin.from("mailing_events").insert({
          campaign_id, subscriber_id: r.id,
          event_type: "failed_to_send",
          metadata: { error: sendRes.error.message },
        });
      } else {
        sent++;
        await admin.from("mailing_events").insert({
          campaign_id, subscriber_id: r.id,
          event_type: "sent",
          resend_event_id: sendRes.data?.id || null,
        });
      }
    }

    await admin.from("mailing_campaigns").update({
      status: failed === (recipients?.length || 0) ? "failed" : "sent",
      recipient_count: sent,
      sent_at: new Date().toISOString(),
    }).eq("id", campaign_id);

    return new Response(
      JSON.stringify({ success: true, sent, failed, total: recipients?.length || 0 }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
