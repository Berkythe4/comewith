-- 254: DI #1 and DI #2 founder money -> Founder sponsorship (Keith, 2026-11-18).
--
-- Same rule as 253: on a fundraiser a DONATION reads as money for the MS Society.
-- Keith's personal money covering production costs is SPONSORSHIP, tier
-- 'Founder sponsorship', kept separate from the event's other sponsors.
--   DI #1  bd7704cf  $1,800.00  "Founder donation covering DI#1 costs"
--   DI #2  36f3266f    $162.44  "Founder donation covering remaining costs"
-- Every view sums donations + sponsor cash into gross / total raised
-- (v_event_summary, v_kpi_event_financials, v_event_money, v_pl_monthly), so totals
-- are unchanged; only the label moves (Donations -> Sponsorship). Asserted below.
-- The public impact reports are stored jsonb and are not touched (DI #1's already
-- calls the $1,800 an expense Keith paid personally, never a donation).

begin;

create temp table _before on commit drop as
  select e.id, coalesce(d.amt,0) + coalesce(s.amt,0) total
    from events e
    left join lateral (select sum(amount) amt from third_party_donations where event_id = e.id) d on true
    left join lateral (select sum(cash_amount) amt from sponsorships where event_id = e.id and status <> 'cancelled') s on true
   where e.id in ('b65d425a-967f-46a2-af57-2ca024f88106', 'ff2b1917-9a3f-42f8-8b9f-4a1deec2338e');

do $$ begin
  if (select count(*) from third_party_donations where id in
        ('bd7704cf-747a-470c-a69b-e04dc9e72703', '36f3266f-3e7b-45ca-8b7d-77470b28c51d')
        and payment_processor = 'Founder out-of-pocket') <> 2
  then raise exception 'founder rows not as expected'; end if;
end $$;

insert into sponsorships (event_id, actor_id, tier, cash_amount, status, notes, payment_refs)
select d.event_id, d.actor_id, 'Founder sponsorship', d.amount, 'paid',
       'Keith''s founder sponsorship: his own money covering production costs. NOT a donation and not money to the MS Society. Moved from third_party_donations ' || d.id || ' (was: "' || d.notes || '") on 2026-11-18 (Keith; same rule as DI #3, 253).',
       array['third_party_donation:' || d.id]
  from third_party_donations d
 where d.id in ('bd7704cf-747a-470c-a69b-e04dc9e72703', '36f3266f-3e7b-45ca-8b7d-77470b28c51d');

delete from third_party_donations
 where id in ('bd7704cf-747a-470c-a69b-e04dc9e72703', '36f3266f-3e7b-45ca-8b7d-77470b28c51d');

do $$ declare bad int; begin
  select count(*) into bad
    from _before b
    left join lateral (select sum(amount) amt from third_party_donations where event_id = b.id) d on true
    left join lateral (select sum(cash_amount) amt from sponsorships where event_id = b.id and status <> 'cancelled') s on true
   where coalesce(d.amt,0) + coalesce(s.amt,0) <> b.total;
  if bad <> 0 then raise exception 'donations + sponsors changed on % event(s)', bad; end if;
  if exists (select 1 from third_party_donations where payment_processor ilike 'founder%')
  then raise exception 'founder donation rows remain'; end if;
end $$;

commit;
