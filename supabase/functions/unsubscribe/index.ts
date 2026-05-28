// unsubscribe
//
// Public endpoint. POST {token} → flips subscribers.status to
// 'unsubscribed', sets unsubscribed_at.
//
// Global unsubscribe per master-list architecture: one click removes
// the subscriber from ALL segments at once. No per-segment unsub.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

function jsonError(s: number, m: string) {
  return new Response(JSON.stringify({ error: m }), { status: s, headers: JSON_HEADERS });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const body = await req.json().catch(() => ({}));
    const token = (body.token || "").toString().trim();
    if (!token) return jsonError(400, "token required");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { data: sub, error: lookupErr } = await admin
      .from("subscribers")
      .select("id, email, status")
      .eq("unsubscribe_token", token)
      .maybeSingle();

    if (lookupErr || !sub) return jsonError(404, "Invalid unsubscribe link");

    if (sub.status === "unsubscribed") {
      return new Response(
        JSON.stringify({ success: true, already: true, email: sub.email }),
        { headers: JSON_HEADERS },
      );
    }

    const now = new Date().toISOString();
    const { error: updErr } = await admin
      .from("subscribers")
      .update({ status: "unsubscribed", unsubscribed_at: now })
      .eq("id", sub.id);

    if (updErr) return jsonError(500, "Could not unsubscribe: " + updErr.message);

    return new Response(
      JSON.stringify({ success: true, already: false, email: sub.email, unsubscribed_at: now }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
