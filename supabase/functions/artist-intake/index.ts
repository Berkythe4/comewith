// artist-intake
//
// PUBLIC (deploy with --no-verify-jwt). A NEW artist fills artist-intake.html and
// this creates them as an `actors` row with the 'artist' role (+ optional photo),
// generating their edit_token so they immediately get a self-service edit link
// (artist-edit.html?token=…). public_profile stays FALSE — an admin reviews and
// toggles them onto the collective from the dashboard Artists tab. Dedupes by
// email so a re-submit returns the existing profile instead of a duplicate actor.
// Mirrors artist-self (service role, publishable key as the gateway apikey).

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
// Fields copied straight through from the form (display_name + email handled separately).
const FIELDS = ["legal_name", "phone", "instagram", "soundcloud", "tiktok", "website", "bio"];

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const b = await req.json().catch(() => ({}));

  // Honeypot: bots fill hidden fields humans never see. Pretend success, create nothing.
  if ((b.hp ?? b.company ?? "").toString().trim()) {
    return new Response(JSON.stringify({ ok: true, display_name: "" }), { headers: JH });
  }

  const display_name = (b.display_name || "").toString().trim();
  if (!display_name) return err(400, "Your artist / DJ name is required.");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const admin = createClient(SUPA, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const SITE_URL = (Deno.env.get("SITE_URL") || "").replace(/\/+$/, "");
  const editLink = (t: string) => `${SITE_URL}/artist-edit.html?token=${t}`;

  const email = (b.email || "").toString().trim().toLowerCase() || null;

  // Dedupe by email — return the existing profile's edit link rather than a duplicate.
  if (email) {
    const { data: existing } = await admin
      .from("actors").select("id, display_name, edit_token")
      .ilike("email", email).is("deleted_at", null).maybeSingle();
    if (existing) {
      const { data: role } = await admin.from("actor_roles")
        .select("id").eq("actor_id", existing.id).eq("role", "artist").maybeSingle();
      if (!role) await admin.from("actor_roles").insert({ actor_id: existing.id, role: "artist" });
      return new Response(JSON.stringify({
        ok: true, already: true, display_name: existing.display_name,
        edit_url: editLink(existing.edit_token),
      }), { headers: JH });
    }
  }

  // Build + insert the actor (public_profile / status / kind / edit_token all default).
  const row: Record<string, string> = { display_name, kind: "person" };
  for (const f of FIELDS) if (f in b) { const v = (b[f] == null ? "" : String(b[f])).trim(); if (v) row[f] = v; }
  if (email) row.email = email;

  const { data: created, error } = await admin.from("actors").insert(row).select("id, edit_token").single();
  if (error || !created) return err(500, error?.message || "Could not create the profile.");
  await admin.from("actor_roles").insert({ actor_id: created.id, role: "artist" });

  // Optional photo (base64 data URL) → event-photos bucket, like artist-self.
  if (b.photo && /^data:image\//i.test(String(b.photo))) {
    const m = String(b.photo).match(/^data:(image\/[a-z+]+);base64,(.+)$/i);
    if (m) {
      const mime = m[1];
      const bytes = Uint8Array.from(atob(m[2]), (c) => c.charCodeAt(0));
      if (bytes.length && bytes.length <= 8 * 1024 * 1024) {
        const ext = mime.includes("png") ? "png" : mime.includes("webp") ? "webp" : "jpg";
        const path = `artist/${created.id}/${Date.now()}_intake.${ext}`;
        const { error: up } = await admin.storage.from("event-photos").upload(path, bytes, { contentType: mime, upsert: true });
        if (!up) await admin.from("actors").update({ photo_path: path }).eq("id", created.id);
      }
    }
  }

  // Non-fatal admin ping so Keith knows to review + publish.
  try {
    const apiKey = Deno.env.get("RESEND_API_KEY");
    if (apiKey) {
      const info = [["Name", display_name], ["Email", email || "—"], ["Instagram", row.instagram || "—"]]
        .map(([k, v]) => `<tr><td style="padding:2px 12px 2px 0"><b>${k}</b></td><td>${v}</td></tr>`).join("");
      await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: { "Authorization": `Bearer ${apiKey}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          from: "Come With <berky@comewith.org>", to: "berky@comewith.org",
          reply_to: "berky@comewith.org",
          subject: `New artist intake: ${display_name}`,
          html: `<p>A new artist submitted the intake form. They're added as an actor (<b>hidden</b>) — review and toggle them onto the collective in the dashboard Artists tab.</p><table>${info}</table>`,
        }),
      });
    }
  } catch (_) { /* ignore */ }

  return new Response(JSON.stringify({
    ok: true, display_name, edit_url: editLink(created.edit_token),
  }), { headers: JH });
});
