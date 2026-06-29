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

// Render an email body: leave real HTML alone, but convert plain-text line breaks
// to <br> so typed paragraphs (e.g. "Hey,\n\n...\n\n— Berky") don't collapse into
// one block. Mirrors renderEmailBody() in dashboard.html (preview == sent).
function renderBody(raw: string): string {
  const s = raw || "";
  if (/<(p|br|div|table|h[1-6]|ul|ol|li|a|strong|em|span|img)\b/i.test(s)) return s;
  const esc = s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5">${esc.replace(/\r\n/g, "\n").replace(/\n/g, "<br>")}</div>`;
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
    // Optional: a single address to send a one-off TEST to (no list send, no status change).
    const test_email = (body.test_email || "").toString().trim();

    // Load campaign
    const { data: campaign, error: cErr } = await admin
      .from("mailing_campaigns")
      .select("id, name, subject, preview_text, body_html, body_text, segment_filter, status, cc, survey_id")
      .eq("id", campaign_id)
      .single();
    if (cErr || !campaign) return jsonError(404, "campaign not found");

    // An attached OPEN survey → each recipient gets a personal tokenized link.
    let survey: any = null;
    if (campaign.survey_id) {
      const { data: sv } = await admin.from("surveys").select("id, event_id, status, public_token").eq("id", campaign.survey_id).maybeSingle();
      if (sv && sv.status === "open") survey = sv;
    }
    const surveyCta = (link: string) => `<p style="text-align:center;margin:22px 0 4px"><a href="${link}" style="display:inline-block;background:#16243f;color:#fff;padding:11px 22px;border-radius:8px;text-decoration:none;font-weight:700">Share your feedback →</a></p>`;
    // Placement control: {{survey}} → the button, {{survey_link}} → the raw URL.
    // No placeholder → the button is appended at the end.
    const injectSurvey = (html: string, link: string) => {
      let out = html, used = false;
      if (/\{\{\s*survey_link\s*\}\}/i.test(out)) { out = out.replace(/\{\{\s*survey_link\s*\}\}/ig, link); used = true; }
      if (/\{\{\s*survey(_button)?\s*\}\}/i.test(out)) { out = out.replace(/\{\{\s*survey(_button)?\s*\}\}/ig, surveyCta(link)); used = true; }
      return used ? out : out + surveyCta(link);
    };

    const resendKey = Deno.env.get("RESEND_API_KEY");
    if (!resendKey) return jsonError(500, "RESEND_API_KEY not set");
    const resend = new Resend(resendKey);

    // TEST SEND — one email to the given address. No list send, no status change,
    // no mailing_events logging. Works on any campaign (even an already-sent one).
    if (test_email) {
      const footer = `<hr style="border:none;border-top:1px solid #ddd;margin:32px 0 12px;"><p style="font-size:0.75rem;color:#8A7F72;text-align:center;">This is a <strong>TEST</strong> send. The real email includes a working personal unsubscribe link.</p>`;
      let bodyHtml = renderBody(campaign.body_html || campaign.body_text || "");
      if (survey) bodyHtml = injectSurvey(bodyHtml, `${SITE_URL}/survey.html?t=${survey.public_token}`);
      const html = bodyHtml + footer;
      const testRes = await resend.emails.send({
        from: FROM,
        to: test_email,
        replyTo: REPLY_TO,
        subject: "[TEST] " + campaign.subject,
        html,
      });
      if (testRes.error) {
        console.error("test send failed:", JSON.stringify(testRes.error));
        return jsonError(502, "Resend rejected the test: " + testRes.error.message);
      }
      return new Response(
        JSON.stringify({ success: true, test: true, sent_to: test_email, id: testRes.data?.id || null }),
        { headers: JSON_HEADERS },
      );
    }

    if (campaign.status === "sent") return jsonError(409, "campaign already sent");
    if (campaign.status === "sending") return jsonError(409, "campaign send in progress");

    // Mark sending
    await admin.from("mailing_campaigns").update({ status: "sending" }).eq("id", campaign_id);

    // Pick subscriber recipients. segment_filter empty = all subscribed; else that
    // segment. An empty segment is NOT an early exit — any CC addresses still send.
    let recipients: any[] = [];
    {
      let recipientQuery = admin
        .from("subscribers")
        .select("id, email, full_name, unsubscribe_token")
        .eq("status", "subscribed");
      let runIt = true;
      if (campaign.segment_filter) {
        const { data: segRows } = await admin
          .from("subscriber_segments")
          .select("subscriber_id")
          .eq("segment", campaign.segment_filter);
        const ids = (segRows || []).map((r) => r.subscriber_id);
        if (ids.length === 0) runIt = false;
        else recipientQuery = recipientQuery.in("id", ids);
      }
      if (runIt) {
        const { data, error: rErr } = await recipientQuery;
        if (rErr) {
          await admin.from("mailing_campaigns").update({ status: "failed" }).eq("id", campaign_id);
          return jsonError(500, "couldn't load recipients: " + rErr.message);
        }
        recipients = data || [];
      }
    }

    let sent = 0;
    let failed = 0;
    for (const r of recipients) {
      const unsubUrl = `${UNSUB_BASE}?token=${r.unsubscribe_token}`;
      const footer = `<hr style="border:none;border-top:1px solid #ddd;margin:32px 0 12px;"><p style="font-size:0.75rem;color:#8A7F72;text-align:center;">You're getting this because you subscribed to Come With updates. <a href="${unsubUrl}" style="color:#8A7F72;">Unsubscribe</a> anytime.</p>`;
      let bodyHtml = renderBody(campaign.body_html || campaign.body_text || "");
      if (survey) {
        const { data: inv } = await admin.from("survey_invites").insert({ survey_id: survey.id, subscriber_id: r.id, event_id: survey.event_id, email: r.email, label: r.full_name }).select("token").single();
        bodyHtml = injectSurvey(bodyHtml, `${SITE_URL}/survey.html?t=${inv ? inv.token : survey.public_token}`);
      }
      const html = bodyHtml + footer;

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

    // CC / "also send to": extra addresses, each gets their own copy. Not subscribers,
    // so no unsubscribe token and no mailing_events row. De-duped vs subscriber emails.
    const subEmails = new Set(recipients.map((r) => (r.email || "").toLowerCase()));
    const ccList = [...new Set(
      (campaign.cc || "")
        .split(/[,;\s]+/)
        .map((s: string) => s.trim().toLowerCase())
        .filter((s: string) => s.includes("@") && !subEmails.has(s)),
    )];
    let ccSent = 0;
    for (const email of ccList) {
      const footer = `<hr style="border:none;border-top:1px solid #ddd;margin:32px 0 12px;"><p style="font-size:0.75rem;color:#8A7F72;text-align:center;">You were CC'd on this Come With update.</p>`;
      let bodyHtml = renderBody(campaign.body_html || campaign.body_text || "");
      if (survey) {
        const { data: inv } = await admin.from("survey_invites").insert({ survey_id: survey.id, event_id: survey.event_id, email }).select("token").single();
        bodyHtml = injectSurvey(bodyHtml, `${SITE_URL}/survey.html?t=${inv ? inv.token : survey.public_token}`);
      }
      const html = bodyHtml + footer;
      const res = await resend.emails.send({ from: FROM, to: email, replyTo: REPLY_TO, subject: campaign.subject, html });
      if (res.error) failed++; else { sent++; ccSent++; }
    }

    await admin.from("mailing_campaigns").update({
      status: sent > 0 ? "sent" : "failed",
      recipient_count: sent,
      sent_at: new Date().toISOString(),
    }).eq("id", campaign_id);

    return new Response(
      JSON.stringify({ success: true, sent, failed, cc: ccSent, total: recipients.length + ccList.length }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
