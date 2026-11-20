-- 257: DI #3 tip guarantee corrected from Signal's sales summary + Keith (2026-11-20);
--      the private 'agreements' bucket accepts spreadsheets so vendor reports can be filed.
--
-- Keith, 2026-11-20: the tip guarantee was $800 (not $500). Guests tipped $480.90 on card
-- (Signal sales summary) + $60 cash = $540.90, so the top-up owed was $259.10. On the night
-- Gavin said $500 was owed; Keith paid $500 by Venmo -> overpaid by $240.90. The 39 voided
-- items ($509) on Signal's report ARE the prepaid drink tickets (Keith confirmed).
--
-- NO AMOUNTS CHANGE. $500 is what actually left Keith's Venmo, so the expense stays $500 and
-- stays inside his founder sponsorship (253). Whether $240.90 is claimed back from Signal is
-- Keith's call; nothing is booked as owed until it is decided. This corrects the WORDS: the
-- expense and budget line both said the $500 "met the guarantee", which is wrong.
--
-- Bucket: 'agreements' is private (public = false) and its storage RLS is unchanged; this only
-- widens which file types it accepts (xlsx, xls, csv) alongside pdf/html.

begin;

do $$
declare n int;
begin
  -- 1. tip guarantee expense (Venmo $500)
  update expenses
     set description = 'Signal tip guarantee for DI #3: guarantee was $800. Guests tipped $540.90 ($480.90 card per Signal''s sales summary + $60 cash), so $259.10 was owed; Gavin claimed $500 on the night and Keith paid $500 by Venmo - $240.90 more than owed (dispute open, see task). Owner-funded; part of Keith''s $1,750 founder sponsorship of DI #3 (253). Corrected 2026-11-20 (Keith).',
         updated_at = now()
   where id = 'e0e94bc1-0e9d-41a1-b00c-7525d521e4f5' and amount = 500 and deleted_at is null;
  get diagnostics n = row_count; if n <> 1 then raise exception '257: tip expense not found (%)', n; end if;

  -- 2. budget line
  update budget_lines
     set notes = 'FINAL 2026-11-20: guarantee $800; guests tipped $540.90 ($480.90 card + $60 cash) so $259.10 was owed. Keith paid $500 (Venmo) at Gavin''s request: $240.90 overpaid, dispute open. Planned amount = what was paid.',
         updated_at = now()
   where id = '8d668563-b937-404b-a759-2be23a82405c' and planned_amount = 500 and deleted_at is null;
  get diagnostics n = row_count; if n <> 1 then raise exception '257: tip budget line not found (%)', n; end if;

  -- 3. the 11/4 Signal payment's description: remembered "~$3,000" -> Signal's figure
  update expenses
     set description = replace(description, 'the $6,000 bar minimum was not met (bar rang ~$3,000)',
                       'the $6,000 bar minimum was not met (bar rang $2,860 net per Signal''s sales summary; the 39 drink tickets rang and were voided, $509, so not counted)'),
         updated_at = now()
   where id = '77b92680-cc46-47aa-8c70-d1d42e65778c' and amount = 3000;
  if (select description from expenses where id = '77b92680-cc46-47aa-8c70-d1d42e65778c') not like '%$2,860 net%' then
    raise exception '257: 11/4 Signal description not updated';
  end if;

  -- 4. tasks from the 2026-11-20 meeting
  update tasks
     set title = 'Claim back the $240.90 tip overpayment from Signal (Gavin)',
         description = 'Guarantee was $800. Guests tipped $540.90 ($480.90 card per Signal''s sales summary, filed on DI #3 + $60 cash), so the top-up owed was $259.10. Gavin claimed $500 on the night; Keith paid $500 by Venmo. Overpaid: $240.90. Keith confirmed the discrepancy with Izzy. If refunded, record it against Keith''s founder sponsorship (the $500 sits inside it).',
         updated_at = now()
   where id = '2afbfbfe-6c3b-4913-83e4-cdb3ec2d6779';
  get diagnostics n = row_count; if n <> 1 then raise exception '257: tip task not found'; end if;

  update tasks
     set title = 'Ask Signal to count the 39 voided drink tickets ($509) toward the minimum',
         description = 'Confirmed (Keith): the 39 voided items ($509) on Signal''s sales summary are our prepaid drink tickets. They are not in the $2,860 net, so the true bar figure is ~$3,369. DI #3''s minimum was missed either way; get it agreed for the DI #4/#5 deal so prepaid drink tickets count.',
         updated_at = now()
   where id = '01fb3847-fc53-495c-8215-4c9e1ca98dc4';
  get diagnostics n = row_count; if n <> 1 then raise exception '257: void task not found'; end if;

  -- 5. DI #3 Control Center note: voids confirmed
  update event_control
     set notes = replace(notes, '39 items / $509 voided, likely the drink tickets', '39 items / $509 voided = the prepaid drink tickets, confirmed by Keith'),
         updated_at = now()
   where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b';
  if (select notes from event_control where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b') not like '%confirmed by Keith%' then
    raise exception '257: event_control note not updated';
  end if;

  -- 6. agreements bucket: accept spreadsheets
  update storage.buckets
     set allowed_mime_types = array['application/pdf', 'text/html',
           'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
           'application/vnd.ms-excel', 'text/csv']
   where id = 'agreements' and public = false;
  get diagnostics n = row_count; if n <> 1 then raise exception '257: agreements bucket not updated'; end if;
end $$;

commit;
