// survey-send — ADMIN (verify_jwt off; checks role manually like send-campaign).
// Creates a tokenized invite per recipient and (optionally) emails each their link
// via Resend. The dashboard builds the recipient list (event guests / segment /
// selected actors) and passes it in explicitly.
import { createClient } from "npm:@supabase/supabase-js@2";
import { Resend } from "npm:resend@4";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const FROM = "Come With <berky@comewith.org>";
const REPLY_TO = "berky@comewith.org";
const SITE_URL = Deno.env.get("SITE_URL") || "https://comewith.org";
const esc = (s: string) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");
  try {
    const auth = req.headers.get("Authorization");
    if (!auth) return err(401, "auth required");
    const userClient = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await userClient.auth.getUser();
    if (!user) return err(401, "invalid session");
    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
    const { data: profile } = await admin.from("profiles").select("role").eq("id", user.id).single();
    if (!profile || !["master_admin", "sub_admin"].includes(profile.role)) return err(403, "admin only");

    const body = await req.json().catch(() => ({}));
    const survey_id = (body.survey_id || "").toString().trim();
    const recipients = Array.isArray(body.recipients) ? body.recipients : [];
    const sendEmail = body.send_email !== false; // default true
    if (!survey_id) return err(400, "survey_id required");
    if (!recipients.length) return err(400, "no recipients");

    const { data: survey } = await admin.from("surveys").select("id, title, intro, status").eq("id", survey_id).single();
    if (!survey) return err(404, "survey not found");

    const resendKey = Deno.env.get("RESEND_API_KEY");
    const resend = resendKey ? new Resend(resendKey) : null;

    let created = 0, sent = 0, failed = 0;
    const links: any[] = [];
    for (const r of recipients) {
      const { data: inv, error } = await admin.from("survey_invites").insert({
        survey_id, event_id: r.event_id || null, actor_id: r.actor_id || null,
        guest_id: r.guest_id || null, subscriber_id: r.subscriber_id || null,
        email: r.email || null, label: r.label || null,
      }).select("id, token").single();
      if (error || !inv) { failed++; continue; }
      created++;
      const link = `${SITE_URL}/survey.html?t=${inv.token}`;
      links.push({ email: r.email || null, label: r.label || null, link });
      if (sendEmail && resend && r.email) {
        const html = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#16243f">
          <p>Hi${r.label ? " " + esc(r.label) : ""},</p>
          <p>${survey.intro ? esc(survey.intro) : "We'd love your quick feedback — it takes a minute."}</p>
          <p><a href="${link}" style="display:inline-block;background:#16243f;color:#fff;padding:11px 20px;border-radius:8px;text-decoration:none;font-weight:700">Take the survey →</a></p>
          <p style="font-size:12px;color:#888">Or paste this link into your browser:<br>${link}</p></div>`;
        const res = await resend.emails.send({ from: FROM, to: r.email, replyTo: REPLY_TO, subject: survey.title || "Your feedback", html });
        if (res.error) { failed++; } else { sent++; await admin.from("survey_invites").update({ sent_at: new Date().toISOString() }).eq("id", inv.id); }
      }
    }
    return new Response(JSON.stringify({ success: true, created, sent, failed, links }), { headers: JH });
  } catch (e) {
    return err(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
