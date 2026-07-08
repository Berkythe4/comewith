// survey-submit — PUBLIC (verify_jwt off). Saves a response + answers for a token
// (invite or public). Tags the response to the event / actor / guest / subscriber
// carried by the invite (or just the event, for an anonymous public link).
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");
  try {
    const body = await req.json().catch(() => ({}));
    const token = (body.token || "").toString().trim();
    const answers = Array.isArray(body.answers) ? body.answers : [];
    if (!token) return err(400, "token required");

    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    const ctx: any = { invite_id: null, event_id: null, actor_id: null, guest_id: null, subscriber_id: null, anonymous: false };
    let surveyId: string | null = null, invite: any = null;

    const { data: inv } = await admin.from("survey_invites")
      .select("id, survey_id, event_id, actor_id, guest_id, subscriber_id").eq("token", token).maybeSingle();
    if (inv) {
      surveyId = inv.survey_id; invite = inv;
      ctx.invite_id = inv.id; ctx.event_id = inv.event_id; ctx.actor_id = inv.actor_id;
      ctx.guest_id = inv.guest_id; ctx.subscriber_id = inv.subscriber_id;
    } else {
      const { data: sv } = await admin.from("surveys").select("id, allow_anonymous, event_id").eq("public_token", token).maybeSingle();
      if (!sv) return err(404, "survey not found");
      if (!sv.allow_anonymous) return err(403, "This survey needs a personal link.");
      surveyId = sv.id; ctx.anonymous = true; ctx.event_id = sv.event_id;
    }

    const { data: survey } = await admin.from("surveys").select("id, status, event_id").eq("id", surveyId).single();
    if (!survey) return err(404, "survey not found");
    if (survey.status !== "open") return err(409, "This survey is closed.");
    if (!ctx.event_id) ctx.event_id = survey.event_id;

    const { data: resp, error: rErr } = await admin.from("survey_responses").insert({
      survey_id: surveyId, invite_id: ctx.invite_id, event_id: ctx.event_id, actor_id: ctx.actor_id,
      guest_id: ctx.guest_id, subscriber_id: ctx.subscriber_id, anonymous: ctx.anonymous, source: "web",
    }).select("id").single();
    if (rErr || !resp) { console.error("survey-submit response insert failed:", rErr?.message); return err(500, "Could not save your response — please try again."); }

    const rows = answers.filter((a: any) => a && a.question_id)
      .map((a: any) => ({ response_id: resp.id, question_id: a.question_id, value: a.value === undefined ? null : a.value }));
    if (rows.length) {
      const { error: aErr } = await admin.from("survey_answers").insert(rows);
      if (aErr) { console.error("survey-submit answers insert failed:", aErr.message); return err(500, "Could not save your answers — please try again."); }
    }
    if (invite) await admin.from("survey_invites").update({ responded_at: new Date().toISOString() }).eq("id", invite.id);

    return new Response(JSON.stringify({ success: true }), { headers: JH });
  } catch (e) {
    console.error("survey-submit unexpected:", e instanceof Error ? e.message : String(e));
    return err(500, "Something went wrong — please try again.");
  }
});
