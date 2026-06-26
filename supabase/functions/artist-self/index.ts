// artist-self
//
// Public, token-gated self-service for artist profiles. An artist opens
// artist-edit.html?token=<edit_token> and this function lets them read + update
// their own bio/socials/photo without a login. The token = actors.edit_token.
// Deployed with --no-verify-jwt (the artist is anonymous); the publishable key
// is still required as the gateway apikey.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const ALLOWED = ["bio", "instagram", "soundcloud", "tiktok", "website"];

function err(s: number, m: string) {
  return new Response(JSON.stringify({ error: m }), { status: s, headers: JH });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const b = await req.json().catch(() => ({}));
  const token = (b.token || "").toString().trim();
  if (!token) return err(400, "token required");

  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const admin = createClient(SUPA, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

  const { data: a } = await admin
    .from("actors")
    .select("id, display_name, bio, instagram, soundcloud, tiktok, website, photo_path, deleted_at")
    .eq("edit_token", token)
    .maybeSingle();
  if (!a || a.deleted_at) return err(404, "This link is invalid or has expired.");

  const photoUrl = (p: string | null) => p ? `${SUPA}/storage/v1/object/public/event-photos/${p}` : null;
  const action = (b.action || "get").toString();

  if (action === "get") {
    return new Response(JSON.stringify({ ok: true, artist: {
      display_name: a.display_name, bio: a.bio, instagram: a.instagram,
      soundcloud: a.soundcloud, tiktok: a.tiktok, website: a.website, photo_url: photoUrl(a.photo_path),
    } }), { headers: JH });
  }

  if (action === "save") {
    const patch: Record<string, string | null> = {};
    for (const f of ALLOWED) if (f in b) { const v = (b[f] == null ? "" : String(b[f])).trim(); patch[f] = v || null; }
    const { error } = await admin.from("actors").update(patch).eq("id", a.id);
    if (error) return err(500, error.message);
    return new Response(JSON.stringify({ ok: true }), { headers: JH });
  }

  if (action === "photo") {
    const dataUrl = (b.dataUrl || "").toString();
    const m = dataUrl.match(/^data:(image\/[a-z+]+);base64,(.+)$/i);
    if (!m) return err(400, "Could not read that image.");
    const mime = m[1];
    if (![...atob(m[2])].length) return err(400, "Empty image.");
    const bytes = Uint8Array.from(atob(m[2]), (c) => c.charCodeAt(0));
    if (bytes.length > 8 * 1024 * 1024) return err(413, "Image too large.");
    const ext = mime.includes("png") ? "png" : mime.includes("webp") ? "webp" : "jpg";
    const path = `artist/${a.id}/${Date.now()}_self.${ext}`;
    const { error: upErr } = await admin.storage.from("event-photos").upload(path, bytes, { contentType: mime, upsert: true });
    if (upErr) return err(500, upErr.message);
    await admin.from("actors").update({ photo_path: path }).eq("id", a.id);
    return new Response(JSON.stringify({ ok: true, photo_url: photoUrl(path) }), { headers: JH });
  }

  if (action === "photo_remove") {
    await admin.from("actors").update({ photo_path: null }).eq("id", a.id);
    return new Response(JSON.stringify({ ok: true, photo_url: null }), { headers: JH });
  }

  return err(400, "unknown action");
});
