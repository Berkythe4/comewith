// staging/guard.js — shared admin gate for the /staging/ review area.
//
// REUSABLE: any staging page gates itself by adding these two lines in <head>
// (the first prevents a content flash, the second runs this guard):
//
//   <script>document.documentElement.style.visibility='hidden'</script>
//   <script type="module" src="/staging/guard.js"></script>
//
// Auth model reuses the EXACT pattern dashboard.html uses (same Supabase project,
// same publishable key, same profiles.role check). One front door: sign in at
// /dashboard.html, the session persists (localStorage, same origin), and every
// /staging/ page sees it. No second password system.
//
// Honest caveat: this is CLIENT-SIDE gating on a static host — good enough to keep
// review-before-publish pages out of casual/public view, NOT real security. Do not
// put genuinely sensitive data as static files here; that stays in Supabase behind RLS.

import { createClient } from 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm';

// PUBLIC values — safe to embed in frontend per the env contract (identical to dashboard.html).
const SUPABASE_URL = 'https://yaytdosxfhcqatmhctzk.supabase.co';
const SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_IkigzWOTU3ZSMK9DxqwwJw_AaQkShCi';

const SIGN_IN_URL = '/dashboard.html';            // the existing front door (shows login when no session)
const ADMIN_ROLES = ['master_admin', 'sub_admin']; // mirrors public.is_admin()

const sb = createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY, {
  auth: { detectSessionInUrl: true, persistSession: true, autoRefreshToken: true },
});

function reveal() { document.documentElement.style.visibility = ''; }

function goSignIn() {
  // Best-effort: remember where we were headed (dashboard doesn't consume this yet).
  try { sessionStorage.setItem('staging_return', location.pathname + location.search); } catch (e) {}
  location.replace(SIGN_IN_URL);
}

function denyNonAdmin(role) {
  reveal();
  document.body.innerHTML =
    '<div style="font-family:DM Sans,system-ui,sans-serif;max-width:520px;margin:18vh auto;padding:24px;text-align:center;color:#16243f">'
    + '<h1 style="font-family:Bebas Neue,Impact,sans-serif;font-weight:400;letter-spacing:.04em;font-size:40px;margin:0 0 8px">STAGING — ADMINS ONLY</h1>'
    + '<p style="color:#5d6678">This area is limited to Come With admins. Your account'
    + (role ? ' (role: ' + role + ')' : '') + ' doesn’t have access.</p>'
    + '<p><a href="' + SIGN_IN_URL + '" style="color:#16243f;font-weight:700">Sign in with an admin account →</a></p>'
    + '</div>';
}

async function guard() {
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { goSignIn(); return; }

  const { data: profile, error } = await sb
    .from('profiles')
    .select('role')
    .eq('id', session.user.id)
    .single();

  if (error || !profile) { goSignIn(); return; }       // can't confirm role → treat as unauthenticated

  if (ADMIN_ROLES.includes(profile.role)) {
    window.__stagingRole = profile.role;                // available to the page if useful
    reveal();
    document.dispatchEvent(new CustomEvent('staging:authed', { detail: { role: profile.role } }));
  } else {
    denyNonAdmin(profile.role);
  }
}

guard();

export { sb };
