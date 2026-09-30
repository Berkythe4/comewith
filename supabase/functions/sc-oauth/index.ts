// sc-oauth  (PUBLIC — deploy with --no-verify-jwt; it's SoundCloud's redirect target)
//
// OAuth 2.1 Authorization-Code + PKCE CALLBACK only. SoundCloud redirects the
// browser here with ?code&state after the user approves. We match `state` to the
// row the admin `sc-connect?action=start` created (which holds the PKCE verifier),
// exchange the code for tokens, store them, and bounce back to the dashboard.
// The redirect URI registered in the SoundCloud app must be EXACTLY this URL.

import { createClient } from "npm:@supabase/supabase-js@2";

Deno.serve(async (req) => {
  const SUPA = Deno.env.get("SUPABASE_URL")!;
  const SITE = (Deno.env.get("SITE_URL") || "https://comewith.org").replace(/\/+$/, "");
  const back = (qs: string) => Response.redirect(`${SITE}/dashboard.html${qs}`, 302);

  const url = new URL(req.url);
  const code = url.searchParams.get("code");
  const state = url.searchParams.get("state");
  const SRK = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const admin = createClient(SUPA, SRK);
  const clientId = Deno.env.get("SC_CLIENT_ID"), clientSecret = Deno.env.get("SC_CLIENT_SECRET");
  const redirectUri = `${SUPA}/functions/v1/sc-oauth`;

  // ── A GUEST DJ exporting an episode to their own account (migration 215). ──
  // Their state lives in sc_dj_exports, not the singleton. The code is exchanged,
  // the token is handed to sc-connect export_as for this one playlist, and then
  // dropped: it is never written anywhere. They land back on their dj.html link.
  if (state) {
    const { data: dx } = await admin.from("sc_dj_exports").select("state, playlist_id, code_verifier, created_at, completed_at")
      .eq("state", state).maybeSingle();
    if (dx) {
      const { data: ep } = await admin.from("sc_playlists").select("dj_token").eq("id", dx.playlist_id).maybeSingle();
      const djBack = (st: string) => Response.redirect(`${SITE}/dj.html?ep=${encodeURIComponent(ep?.dj_token || "")}&sc=${st}`, 302);
      const finish = async (patch: Record<string, unknown>, st: string) => {
        await admin.from("sc_dj_exports").update({ ...patch, code_verifier: null, completed_at: new Date().toISOString() }).eq("state", state);
        return djBack(st);
      };
      if (!ep?.dj_token) return finish({ ok: false, error: "The DJ link was revoked." }, "error");
      if (url.searchParams.get("error")) return finish({ ok: false, error: "Declined on SoundCloud." }, "denied");
      // One use, and only for 30 minutes: a replayed or stale callback does nothing.
      const fresh = Date.now() - new Date(dx.created_at).getTime() < 30 * 60000;
      if (!code || !dx.code_verifier || dx.completed_at || !fresh) return finish({ ok: false, error: "That SoundCloud approval expired - start the export again." }, "error");
      if (!clientId || !clientSecret) return finish({ ok: false, error: "SoundCloud is not configured." }, "error");
      try {
        const tokenRes = await fetch("https://secure.soundcloud.com/oauth/token", {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
          body: new URLSearchParams({
            grant_type: "authorization_code", client_id: clientId, client_secret: clientSecret,
            redirect_uri: redirectUri, code_verifier: dx.code_verifier, code,
          }),
        });
        const tok = await tokenRes.json().catch(() => ({}));
        if (!tokenRes.ok || !tok.access_token) return finish({ ok: false, error: "SoundCloud did not accept the approval." }, "error");
        let username: string | null = null;
        try {
          const me = await (await fetch("https://api.soundcloud.com/me", { headers: { "Authorization": "OAuth " + tok.access_token, "accept": "application/json; charset=utf-8" } })).json();
          username = me?.username || null;
        } catch { /* non-fatal */ }
        const r = await fetch(`${SUPA}/functions/v1/sc-connect`, {
          method: "POST",
          headers: { "Content-Type": "application/json", "Authorization": "Bearer " + SRK },
          body: JSON.stringify({ action: "export_as", playlist_id: dx.playlist_id, access_token: tok.access_token }),
        });
        const j = await r.json().catch(() => ({}));
        if (!r.ok || !j.success) return finish({ ok: false, sc_username: username, error: (j.error || "SoundCloud rejected the playlist.").toString().slice(0, 300) }, "error");
        return finish({
          ok: true, sc_username: username, result_url: j.url || null, tracks: j.tracks ?? null,
          skipped: { blocked: j.skipped || [], not_on_soundcloud: j.not_on_soundcloud || [] },
        }, "exported");
      } catch (e) {
        console.error("sc-oauth dj:", e instanceof Error ? e.message : String(e));
        return finish({ ok: false, error: "Something went wrong talking to SoundCloud." }, "error");
      }
    }
  }

  // ── Keith's own connection (the dashboard's singleton). ──
  if (url.searchParams.get("error")) return back("?sc=denied");
  if (!code || !state) return back("?sc=error");
  const { data: row } = await admin.from("sc_oauth").select("id, state, code_verifier").eq("id", "singleton").maybeSingle();
  if (!row || row.state !== state || !row.code_verifier) return back("?sc=error");
  if (!clientId || !clientSecret) return back("?sc=notconfigured");

  try {
    const tokenRes = await fetch("https://secure.soundcloud.com/oauth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": "application/json; charset=utf-8" },
      body: new URLSearchParams({
        grant_type: "authorization_code", client_id: clientId, client_secret: clientSecret,
        redirect_uri: redirectUri, code_verifier: row.code_verifier, code,
      }),
    });
    const tok = await tokenRes.json();
    if (!tokenRes.ok || !tok.access_token) { console.error("sc token exchange:", JSON.stringify(tok).slice(0, 200)); return back("?sc=error"); }

    let username: string | null = null, uid: string | null = null;
    try {
      const me = await (await fetch("https://api.soundcloud.com/me", { headers: { "Authorization": "OAuth " + tok.access_token, "accept": "application/json; charset=utf-8" } })).json();
      username = me?.username || null; uid = me?.id != null ? String(me.id) : null;
    } catch { /* non-fatal */ }

    await admin.from("sc_oauth").update({
      access_token: tok.access_token, refresh_token: tok.refresh_token || null,
      expires_at: new Date(Date.now() + (tok.expires_in || 3600) * 1000).toISOString(),
      sc_username: username, sc_user_id: uid, connected_at: new Date().toISOString(),
      state: null, code_verifier: null, updated_at: new Date().toISOString(),
    }).eq("id", "singleton");

    return back("?sc=connected");
  } catch (e) {
    console.error("sc-oauth:", e instanceof Error ? e.message : String(e));
    return back("?sc=error");
  }
});
