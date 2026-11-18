-- 252: DI #3 - Meta ad was Come With's, and Emmy Adelle is paid (Keith, 2026-11-18).
--
-- * The DI #3 Meta ad ($150) is the 10/16 "FACEBK *AMBJF7ERA4, Wilmington, DE" charge on
--   Come With's Bluevine (Keith: "same ad"). The 226 row (Meta, $150, "paid by Keith
--   personally") is therefore a DUPLICATE and is removed. The real charge stays on
--   Come With's books (Come With paid it) and is tagged to DI #3, and Come With's paying
--   it is booked as an in-kind donation to DI #3 (Keith: "a donation from Come With").
--   Processor is 'Founder out-of-pocket (Come With, in kind)': Come With is Keith's
--   business, and the NON_MISSION_DONATION rule (/founder|out-of-pocket/) must keep this
--   out of "% to the mission" - it paid for an ad, it is not mission money.
-- * Emmy Adelle's $800 (invoice #2026-1114) was paid today from DI Bluevine (ACH
--   initiated by Keith). DI Bluevine balance after it: $798.05 (= $1,598.05 on 11/16
--   in the export - $800).

begin;

do $$ begin
  if (select count(*) from expenses where id = 'c432062b-b33a-4c41-bd8b-4c5365f7ed66' and deleted_at is null and amount = 150) <> 1
  then raise exception 'Meta dup row not as expected'; end if;
  if (select count(*) from expenses where id = '43c972fa-5dc7-45fa-b5ac-b32421ef80a4' and event_id is null
        and ledger = 'come_with' and amount = 150 and deleted_at is null) <> 1
  then raise exception '10/16 Facebook row not as expected'; end if;
  if (select count(*) from expenses where id = '6499c779-2a27-4d33-8340-1cee31dcfc66' and status = 'accrued' and amount = 800) <> 1
  then raise exception 'Emmy row not as expected'; end if;
end $$;

update expenses
   set deleted_at = now(),
       description = description || ' | REMOVED 2026-11-18: duplicate. This ad is the 10/16 FACEBK *AMBJF7ERA4 $150 charge on Come With''s Bluevine (expense 43c972fa), not a personal payment (Keith: same ad).'
 where id = 'c432062b-b33a-4c41-bd8b-4c5365f7ed66';

update expenses
   set event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b',
       description = 'Meta Ads for DI #3, paid from Come With''s Bluevine (stays on Come With''s books). Donated to DI #3 by Come With - see third_party_donations ''Founder out-of-pocket (Come With, in kind)''. Keith 2026-11-18.'
 where id = '43c972fa-5dc7-45fa-b5ac-b32421ef80a4';

insert into third_party_donations (event_id, donor_name, amount, payment_processor, date, notes)
values ('f13bf700-7383-497b-91b8-d1f6781e077b', 'Come With', 150, 'Founder out-of-pocket (Come With, in kind)', '2026-10-16',
        'In-kind donation from Come With: it paid the $150 DI #3 Meta ad from its own Bluevine on 10/16 (expense 43c972fa). Keith 2026-11-18. Counts in money raised, not in % to the mission.');

update expenses
   set status = 'paid', settled_at = now(), cash_source = 'bank', payment_method = 'Bluevine (DI)',
       description = description || ' | PAID 2026-11-18 from DI Bluevine (ACH initiated by Keith). DI balance after: $798.05.'
 where id = '6499c779-2a27-4d33-8340-1cee31dcfc66';

do $$ declare c numeric; u numeric; d numeric; begin
  select sum(amount), sum(amount) filter (where status <> 'paid') into c, u from expenses
   where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b' and deleted_at is null;
  select sum(amount) into d from third_party_donations where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';
  if c <> 4950 or u <> 500 or d <> 2240 then
    raise exception 'costs % unpaid % donations %, expected 4950 / 500 / 2240', c, u, d; end if;
end $$;

commit;
