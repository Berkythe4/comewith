// survey-get — PUBLIC (verify_jwt off). Resolves a token (a per-recipient invite
// token OR a survey public_token) to the survey + its questions + recipient context.
// Only returns surveys whose status = 'open'. No table is exposed to anon — this runs
// with the service role, like get-agreement-by-token / get-event-hub.
import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  try {
    const url = new URL(req.url);
    let token = url.searchParams.get("token") || "";
    if (!token && req.method === "POST") { const b = await req.json().catch(() => ({})); token = (b.token || "").toString(); }
    token = token.trim();
    if (!token) return err(400, "token required");

    const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

    const ctx: any = { invite_id: null, event_id: null, actor_id: null, guest_id: null, subscriber_id: null, who: null, anonymous: false };
    let surveyId: string | null = null;

    const { data: inv } = await admin.from("survey_invites")
      .select("id, survey_id, event_id, actor_id, guest_id, subscriber_id, label").eq("token", token).maybeSingle();
    if (inv) {
      surveyId = inv.survey_id;
      ctx.invite_id = inv.id; ctx.event_id = inv.event_id; ctx.actor_id = inv.actor_id;
      ctx.guest_id = inv.guest_id; ctx.subscriber_id = inv.subscriber_id; ctx.who = inv.label;
    } else {
      const { data: sv } = await admin.from("surveys").select("id, allow_anonymous").eq("public_token", token).maybeSingle();
      if (!sv) return err(404, "survey not found");
      if (!sv.allow_anonymous) return err(403, "This survey needs a personal link.");
      surveyId = sv.id; ctx.anonymous = true;
    }

    const { data: survey } = await admin.from("surveys").select("id, title, intro, status, event_id").eq("id", surveyId).single();
    if (!survey) return err(404, "survey not found");
    if (survey.status !== "open") return err(409, "This survey is closed.");
    if (!ctx.event_id) ctx.event_id = survey.event_id;

    let eventName: string | null = null;
    if (ctx.event_id) { const { data: e } = await admin.from("events").select("name").eq("id", ctx.event_id).maybeSingle(); eventName = e?.name || null; }

    const { data: questions } = await admin.from("survey_questions")
      .select("id, sort_order, prompt, qtype, options, required").eq("survey_id", surveyId).order("sort_order");

    return new Response(JSON.stringify({
      survey: { id: survey.id, title: survey.title, intro: survey.intro, event_name: eventName },
      questions: questions || [], context: ctx,
    }), { headers: JH });
  } catch (e) {
    console.error("survey-get unexpected:", e instanceof Error ? e.message : String(e));
    return err(500, "Could not load the survey — please try again.");
  }
});
