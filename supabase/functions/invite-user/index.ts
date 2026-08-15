// invite-user
//
// Master-admin-only endpoint behind the dashboard Team tab. Invites a new
// teammate by email (Supabase sends the invite; they set their own password),
// then sets their role + staff_role on the auto-created profiles row.
//
// Security: the CALLER's JWT is verified to belong to a master_admin before any
// admin action runs. Only then is the service-role key used to invite.
//
// NOT YET DEPLOYED. Deploy with:
//   SUPABASE_ACCESS_TOKEN=$SBP_PAT supabase functions deploy invite-user \
//     --project-ref yaytdosxfhcqatmhctzk
//
// IMPORTANT: until migration 043 (financial gate) is applied, a new sub_admin
// can read the financial views via direct REST (the dashboard UI hides Finance
// from non-finance roles, but RLS does not yet). Prefer applying 043 before
// inviting staff who should not see financials.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

function jsonError(status: number, message: string) {
  return new Response(JSON.stringify({ error: message }), { status, headers: JSON_HEADERS });
}

const VALID_ROLES = ["sub_admin", "master_admin"];
const VALID_STAFF = ["operations", "marketing", "full"];

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const url = Deno.env.get("SUPABASE_URL")!;
    const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const anonKey = Deno.env.get("SUPABASE_ANON_KEY")!;

    // 1. Verify the caller is a master_admin (using their own JWT).
    const authHeader = req.headers.get("Authorization") || "";
    if (!authHeader) return jsonError(401, "Missing Authorization header");
    const caller = createClient(url, anonKey, { global: { headers: { Authorization: authHeader } } });
    const { data: { user }, error: userErr } = await caller.auth.getUser();
    if (userErr || !user) return jsonError(401, "Invalid session");

    const admin = createClient(url, serviceKey);
    const { data: callerProfile } = await admin.from("profiles").select("role").eq("id", user.id).single();
    if (callerProfile?.role !== "master_admin") return jsonError(403, "Master admin only");

    // 2. Validate input.
    const b = await req.json().catch(() => ({}));
    const email = (b.email || "").toString().trim().toLowerCase();
    const full_name = (b.full_name || "").toString().trim() || null;
    const role = VALID_ROLES.includes(b.role) ? b.role : "sub_admin";
    const staff_role = VALID_STAFF.includes(b.staff_role) ? b.staff_role : null;
    if (!email) return jsonError(400, "email required");

    // The site-owner guard (138) is a trigger on profiles, and it deliberately
    // exempts service-role callers so break-glass repair stays possible. This
    // function IS a service-role caller, so it has to enforce the rule itself:
    // an invite aimed at the owner's address must never reach step 4's profile
    // patch. Supabase already errors on an existing email, but that is its
    // behaviour, not our guarantee.
    const { data: owner } = await admin.from("profiles").select("email").eq("is_owner", true).maybeSingle();
    if (owner?.email && owner.email.toLowerCase() === email) {
      return jsonError(403, "That address belongs to the site owner and can't be re-invited.");
    }

    // 3. Invite (creates auth user; the handle_new_user trigger creates the
    //    profiles row). Pass full_name through user metadata.
    const { data: invited, error: inviteErr } = await admin.auth.admin.inviteUserByEmail(email, {
      data: full_name ? { full_name } : undefined,
    });
    if (inviteErr) return jsonError(400, "Invite failed: " + inviteErr.message);
    const newId = invited.user?.id;
    if (!newId) return jsonError(500, "Invite returned no user id");

    // 4. Set role + staff_role + full_name on the new profile.
    const patch: Record<string, unknown> = { role, staff_role };
    if (full_name) patch.full_name = full_name;
    const { error: updErr } = await admin.from("profiles").update(patch).eq("id", newId);
    if (updErr) return jsonError(500, "Profile update failed: " + updErr.message);

    return new Response(
      JSON.stringify({ success: true, user_id: newId, email, role, staff_role }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
