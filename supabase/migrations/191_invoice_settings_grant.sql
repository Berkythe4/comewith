-- ============================================================
-- COME WITH — 191 Payment details could not be opened by anyone
--
-- 188 ended with:
--
--   revoke all on public.invoice_settings from anon, authenticated;
--   grant  all on public.invoice_settings to   service_role;
--
-- with a comment claiming "authenticated reaches invoice_settings only through
-- the edge function". That was true of the invoice RENDERER, which runs as the
-- service role — and false of the settings SCREEN, which is ordinary dashboard
-- code reading `sb.from('invoice_settings')` as the signed-in user. So the
-- screen was unreachable for everybody, including the owner.
--
-- THE SHAPE OF THE MISTAKE, which is worth more than the fix: GRANTS ARE
-- CHECKED BEFORE RLS. Revoking the grant does not make a table "admin only", it
-- makes it nobody-only, and the failure looks exactly like a permission problem
-- with the CALLER rather than with the table. The dashboard reported it as
-- "master-admin only" and sent Keith to check his own role, which was correct
-- all along.
--
-- The right shape is the one every other admin table in this database uses:
-- grant the privilege to `authenticated`, and let RLS decide who among them.
-- The policy from 188 is unchanged and still master-admin only, so Janelle and
-- Liz still cannot read the account number, and anon still holds nothing.
--
-- SELECT and UPDATE only. It is one fixed row (id = true, enforced by CHECK);
-- nobody should be inserting a second set of payment details or deleting the
-- only one.
-- ============================================================
begin;

grant select, update on public.invoice_settings to authenticated;

-- anon keeps nothing at all. Re-asserted rather than assumed: 013's default
-- privileges are exactly the mechanism that re-granted views in the 016/017
-- regression, so every migration that touches grants here says it out loud.
revoke all on public.invoice_settings from anon;

commit;

-- DOWN: revoke select, update on public.invoice_settings from authenticated;
--       (which restores the 188 bug — the screen goes dead again.)
