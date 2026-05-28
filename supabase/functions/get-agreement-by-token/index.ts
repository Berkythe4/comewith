// get-agreement-by-token
//
// Public endpoint called by sign.html. Validates a token (must exist,
// not be expired, not have been used), then returns the agreement
// fields the customer needs to review and sign.
//
// Does NOT include internal admin fields (notes, created_by, etc.).
// Service-role client bypasses RLS for the read.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

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
    if (!token || typeof token !== "string") return jsonError(400, "token is required");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    const { data: link, error: linkErr } = await admin
      .from("agreement_links")
      .select("agreement_id, expires_at, used_at")
      .eq("token", token)
      .single();

    if (linkErr || !link) return jsonError(404, "Invalid signing link");
    if (new Date(link.expires_at) < new Date()) {
      return jsonError(410, "This signing link has expired. Ask Berky for a fresh one.");
    }
    const alreadySigned = !!link.used_at;

    const { data: agreement, error: agErr } = await admin
      .from("agreements")
      .select(
        "id, agreement_type, status, event_date, event_start_time, event_end_time, venue_name, venue_address, services, equipment, subtotal, deposit_amount, total_amount, payment_method, payment_notes, recording_rights, promo_rights, rental_start, rental_return, client_signed_at, client_signature_url, client:clients(full_name, email)",
      )
      .eq("id", link.agreement_id)
      .single();

    if (agErr || !agreement) return jsonError(404, "Agreement not found");

    return new Response(
      JSON.stringify({
        agreement,
        link: {
          expires_at: link.expires_at,
          already_signed: alreadySigned,
        },
      }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
