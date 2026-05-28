// confirm-subscription
//
// Public endpoint called by confirm.html (the link in the
// double-opt-in email). POST {token} → flips subscribers.status
// from 'pending' to 'subscribed', sets confirmed_at.
// Returns subscriber email (truncated for privacy) and segment count.

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
      .select("id, email, status, confirmed_at")
      .eq("unsubscribe_token", token)
      .maybeSingle();

    if (lookupErr || !sub) return jsonError(404, "Invalid confirmation link");

    if (sub.status === "subscribed") {
      return new Response(
        JSON.stringify({ success: true, already: true, email: sub.email }),
        { headers: JSON_HEADERS },
      );
    }
    if (sub.status !== "pending" && sub.status !== "unsubscribed") {
      return jsonError(409, "This subscription can't be confirmed (status: " + sub.status + ")");
    }

    const now = new Date().toISOString();
    const { error: updErr } = await admin
      .from("subscribers")
      .update({ status: "subscribed", confirmed_at: now })
      .eq("id", sub.id);

    if (updErr) return jsonError(500, "Could not confirm: " + updErr.message);

    return new Response(
      JSON.stringify({ success: true, already: false, email: sub.email, confirmed_at: now }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
