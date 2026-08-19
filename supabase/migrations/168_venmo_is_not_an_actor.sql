-- ============================================================
-- COME WITH — 168 a payment rail is not a counterparty
--
-- 158 seeded 'Venmo' as an actor with an alias rule matching 'venmo'. Every
-- Venmo payment therefore resolves to a single payee called Venmo, no matter who
-- actually received the money. 167 renamed the two charges apart and the
-- grouping did not move, because the actor link outranks the vendor text.
--
-- This is a bank-statement descriptor mistaken for a person. It is the same
-- error as filing every card charge under 'Visa'. Left alone it silently merges
-- unrelated recipients, which is how a $650 false positive appeared on the 1099
-- review list — and, more dangerously, how one person's real total could later
-- hide inside a bucket named after an app.
--
-- No other rail has the problem: PayPal, Zelle, Cash App, Stripe and Square have
-- no actor rows. Checked, not assumed.
--
-- The two charges keep their descriptive vendor text from 167 and simply stop
-- pointing at the fake actor. One recipient is still unnamed; it stays visible
-- as 'Photographer — name needed' rather than being tidied out of sight.
-- ============================================================
begin;

-- Only two expense rows and the one alias rule point at it — every actor_id
-- column in the schema was checked, not guessed.
update public.expenses
   set vendor_actor_id = null
 where vendor_actor_id = (select id from public.actors where display_name = 'Venmo');

-- Stop the next import from recreating the link.
delete from public.vendor_aliases
 where actor_id = (select id from public.actors where display_name = 'Venmo')
    or lower(pattern) = 'venmo';

-- Soft delete, consistent with how every other retired actor is handled here.
update public.actors
   set deleted_at = now()
 where display_name = 'Venmo' and deleted_at is null;

commit;

-- DOWN: clear deleted_at on the actor, re-insert the alias, and relink the two
--   2026 charges. Not worth automating — this record should not come back.
