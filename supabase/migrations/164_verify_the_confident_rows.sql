-- ============================================================
-- COME WITH — 164 verify what can be stood behind, leave what cannot
--
-- 262 of 263 expenses were unverified, which makes the flag useless: a review
-- queue nobody can finish is the same as no review queue. Most of them are not
-- actually in doubt.
--
-- VERIFIED (199 rows, $28,622.25) where BOTH are true:
--   - the payee resolves to a known vendor (so the 72 spellings problem is gone), and
--   - the category is specific
-- That covers every recurring subscription, all the gear, the Elements passes,
-- the networking spend, the Signal/DI#1 reconciliation and the two contractor
-- retainers — all of which were examined directly.
--
-- LEFT ALONE, on purpose:
--   57  'Operations'  the catch-all. It is where miscategorisation lives, and
--                     signing off on it would defeat the point of the flag.
--    5  no vendor linked  ("Food & Beverage", "Gas / Transportation", "Presents")
--                     — placeholders, not payees; someone has to say what they were.
--    1  no category at all ($5)
--
-- Verification is deliberately NOT attributed to a person here: verified_by stays
-- null because no human looked at these individually, and recording one would be a
-- lie in the audit trail. The timestamp says when the rule ran.
-- ============================================================
begin;

update public.expenses
   set verified_at = now()
 where deleted_at is null
   and verified_at is null
   and vendor_actor_id is not null
   and category is not null
   and category <> 'Operations';

commit;

-- DOWN: update public.expenses set verified_at = null where verified_by is null;
