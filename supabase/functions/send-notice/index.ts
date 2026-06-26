// send-notice
//
// Sends a one-off transactional email via Resend (e.g. account / login notices).
// Unlike send-actor-email it does NOT log to conversations — use it for things
// that shouldn't sit in a team-visible thread (password/login instructions).
//
// Auth: service-role bearer (internal/admin scripts) OR a master/sub admin JWT.
// Secret: RESEND_API_KEY.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
const FROM = "Come With <berky@comewith.org>";
const REPLY_TO = "berky@comewith.org";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const URL = Deno.env.get("SUPABASE_URL")!;
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const apiKey = Deno.env.get("RESEND_API_KEY");
  if (!apiKey) return err(500, "RESEND_API_KEY not set");

  // auth: service-role bearer OR admin JWT
  const auth = req.headers.get("Authorization") || "";
  const bearer = auth.replace(/^Bearer\s+/i, "");
  let authed = bearer === SRK;
  if (!authed && bearer) {
    const userClient = createClient(URL, Deno.env.get("SUPABASE_ANON_KEY")!, { global: { headers: { Authorization: auth } } });
    const { data: { user } } = await userClient.auth.getUser();
    if (user) {
      const a = createClient(URL, SRK);
      const { data: prof } = await a.from("profiles").select("role").eq("id", user.id).single();
      authed = !!prof && ["master_admin", "sub_admin"].includes(prof.role);
    }
  }
  if (!authed) return err(401, "unauthorized");

  const b = await req.json().catch(() => ({}));
  const to = b.to, subject = (b.subject || "").toString(), html = (b.html || "").toString();
  if (!to || !subject || !html) return err(400, "to, subject, html required");
  const payload: Record<string, unknown> = { from: FROM, to, reply_to: REPLY_TO, subject, html };
  if (b.cc) payload.cc = b.cc;   // optional carbon-copy

  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { "Authorization": `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const j = await res.json().catch(() => ({}));
  if (!res.ok) return err(502, "Resend: " + (j.message || res.status));
  return new Response(JSON.stringify({ ok: true, id: j.id, to }), { headers: JH });
});
