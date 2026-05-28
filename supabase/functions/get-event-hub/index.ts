// get-event-hub
//
// Public endpoint. POST {slug} returns the assembled public event hub:
//   - event details
//   - sponsors (via sponsorships, sorted by tier/cash)
//   - artists (via artist_bookings, sorted by set_start)
//   - raffle prizes
//   - venue details
//
// Uses service-role to bypass RLS on sponsors/artists/sponsorships/raffle_prizes
// which are admin-only by policy but should be publicly displayable for
// non-cancelled events.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

function jsonError(status: number, message: string) {
  return new Response(JSON.stringify({ error: message }), { status, headers: JSON_HEADERS });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const body = await req.json().catch(() => ({}));
    const slug = (body.slug || "").toString().trim();
    if (!slug) return jsonError(400, "slug is required");

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // Event + venue
    const { data: event, error: eventErr } = await admin
      .from("events")
      .select("id, slug, name, series, event_date, doors_time, end_time, status, bar_minimum, ticket_url, description, hero_image_path, total_attendance, venue:venues(name, city, state, address)")
      .eq("slug", slug)
      .neq("status", "cancelled")
      .single();

    if (eventErr || !event) return jsonError(404, "Event not found");

    // Sponsors via sponsorships (sorted by cash desc as a proxy for tier importance)
    const { data: sponsorships } = await admin
      .from("sponsorships")
      .select("tier, cash_amount, in_kind_value, sponsor:sponsors(name, website, logo_path)")
      .eq("event_id", event.id)
      .neq("status", "cancelled")
      .order("cash_amount", { ascending: false });

    // Artists via artist_bookings
    const { data: bookings } = await admin
      .from("artist_bookings")
      .select("role, set_start, set_end, artist:artists(stage_name, bio, signature_color, photo_path, social_links)")
      .eq("event_id", event.id)
      .order("set_start", { ascending: true, nullsFirst: false });

    // Raffle prizes
    const { data: prizes } = await admin
      .from("raffle_prizes")
      .select("prize_name, donor_name, estimated_value, winner_name")
      .eq("event_id", event.id)
      .order("estimated_value", { ascending: false, nullsFirst: false });

    return new Response(
      JSON.stringify({ event, sponsorships: sponsorships || [], artists: bookings || [], prizes: prizes || [] }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
