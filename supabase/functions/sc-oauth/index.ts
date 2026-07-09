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
  if (url.searchParams.get("error")) return back("?sc=denied");
  if (!code || !state) return back("?sc=error");

  const admin = createClient(SUPA, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const { data: row } = await admin.from("sc_oauth").select("id, state, code_verifier").eq("id", "singleton").maybeSingle();
  if (!row || row.state !== state || !row.code_verifier) return back("?sc=error");

  const clientId = Deno.env.get("SC_CLIENT_ID"), clientSecret = Deno.env.get("SC_CLIENT_SECRET");
  if (!clientId || !clientSecret) return back("?sc=notconfigured");
  const redirectUri = `${SUPA}/functions/v1/sc-oauth`;

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
