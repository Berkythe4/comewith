-- 256: DI #4 / DI #5 placeholders + the 2026-11-20 Dance Infusion committee meeting (Keith).
--
-- Source: "Dance Infusion Committee Meeting - 2026_11_20 10_58 EST - Notes by Gemini.docx".
-- Data only, additive. Nothing here touches money: the $5,000 minimum is a PROPOSAL to Signal,
-- so bar_minimum stays null until Kim agrees; targets live in notes, not kpi_targets (those are
-- edited from the dashboard, never SQL).
--
--  * events: Dance Infusion #4 (Sat 2027-03-06) and #5 (Sat 2027-10-09), Signal, stage/status
--    'planning', not public. Dates are pending Kim's confirmation (task below).
--  * event_control: fundraising on, beneficiary National MS Society (as DI #3).
--  * event_participants: Emmy Adelle as resident headliner on both (agreed in the meeting).
--  * meetings + meeting_notes: the meeting, its 4 decisions, key notes, and one 'action' note
--    per task, linked by task_id.
--  * tasks (source 'meeting'): the 13 Gemini next steps + 4 follow-ups (AOM title sponsor,
--    Crossroads brunch quote, Signal tip overcharge, voided drink tickets). 14 on DI #4, 3 on DI #3.
--  * DI #3 event_control.notes: "~$3,000" -> Signal's sales summary ($2,860 net). Notes only.
--    Each assigned (role 'doer') to the actor linked to that person's login.

begin;

do $$
declare
  v_signal  uuid := 'e417b79d-eb72-4fd7-8182-59dad4a455d9';
  v_di3     uuid := 'f13bf700-7383-497b-91b8-d1f6781e077b';
  v_keith_p uuid := 'bd71f88d-fd0a-4e3a-a712-7ac958318c8b';
  v_henry_p uuid := 'b889cbcd-0a2d-4abc-8d84-bff4b788dd5e';
  v_martin_p uuid := '090ceac3-a8d9-4842-b84d-f0d821bd3b03';
  v_keith   uuid := '09939421-2095-4151-b943-2d943bf83615';  -- actor Berky
  v_henry   uuid := 'e97045d6-aa0b-451c-8711-bc7ef1dfa778';  -- actor Henry
  v_martin  uuid := '400d9160-327d-461a-bf26-1cd9a9723372';  -- actor KRNeY (Martin)
  v_emmy    uuid := 'b7abac4e-9409-4812-b0f3-eb1481e5190c';
  v_di4 uuid; v_di5 uuid; v_mtg uuid; v_task uuid; v_ev uuid;
  r record; i int := 0;
  targets text := 'Targets set 2026-11-20 committee meeting: at least $3,000 in sponsorship and 200 attendees per event (to cover the bar minimum comfortably). Single indoor space; the outdoor tent only pays off at ~250. Saturday, avoiding holidays. Emmy Adelle returns as resident headliner. Proposed to Signal: 2-event commitment (DI #4 + #5) at a $5,000 minimum each - not agreed yet, so bar_minimum is blank.';
begin
  -- guard: nothing named DI #4/#5 exists yet
  if exists (select 1 from events where slug in ('dance-infusion-4','dance-infusion-5')
             or name ilike 'dance infusion #4%' or name ilike 'dance infusion #5%') then
    raise exception '256: DI #4/#5 already exist';
  end if;

  insert into events (slug, name, series, type, event_date, doors_time, end_time, venue_id,
                      capacity, status, stage, is_public, notes)
  values ('dance-infusion-4', 'Dance Infusion #4', 'Dance Infusion', 'dance_infusion', '2027-03-06',
          '15:00', '22:00', v_signal, 400, 'planning', 'planning', false,
          'Date pending Kim (Signal). ' || targets)
  returning id into v_di4;

  insert into events (slug, name, series, type, event_date, doors_time, end_time, venue_id,
                      capacity, status, stage, is_public, notes)
  values ('dance-infusion-5', 'Dance Infusion #5', 'Dance Infusion', 'dance_infusion', '2027-10-09',
          '15:00', '22:00', v_signal, 400, 'planning', 'planning', false,
          'Placeholder. Date pending Kim (Signal). ' || targets)
  returning id into v_di5;

  insert into event_control (event_id, fundraising, beneficiary)
  values (v_di4, true, 'National MS Society'), (v_di5, true, 'National MS Society');

  insert into event_participants (event_id, actor_id, role, roles, notes)
  values (v_di4, v_emmy, 'dj', array['dj'], 'Resident headliner (agreed 2026-11-20 committee meeting)'),
         (v_di5, v_emmy, 'dj', array['dj'], 'Resident headliner (agreed 2026-11-20 committee meeting)');

  insert into meetings (title, meeting_date, attendees, attendee_ids, summary, event_id, status,
                        created_by, visibility)
  values ('Dance Infusion Committee Meeting', '2026-11-20',
          'Keith, Henry, Martin (Janelle, Liz invited)',
          array[v_keith_p, v_henry_p, v_martin_p],
          'DI #3 debrief and next cycle. Dates set for Mar 6 and Oct 9, 2027 (Saturdays, no holidays). Targets: $3,000 sponsorship + 200 attendees per event. Paid ads dropped in favour of local-artist cross-promotion and public health student capstone projects. Source: Gemini notes, reviewed by Claude.',
          v_di4, 'closed', v_keith_p, 'team')
  returning id into v_mtg;

  insert into meeting_notes (meeting_id, kind, body, sort, created_by) values
  (v_mtg, 'decision', 'Every event needs at least $3,000 in sponsorship and 200 attendees.', 1, v_keith_p),
  (v_mtg, 'decision', 'Next two events: Saturday March 6 and Saturday October 9, 2027 (avoids holidays, summer festivals and football season).', 2, v_keith_p),
  (v_mtg, 'decision', 'Partner with local public health school students (capstone projects) for marketing, research and sponsorship groundwork.', 3, v_keith_p),
  (v_mtg, 'decision', 'Keep everyone in the main indoor space; the outdoor tent needs ~250 attendees to justify a second stage.', 4, v_keith_p),
  (v_mtg, 'note', 'DI #3: turnout hurt by competing Friendsgivings; early dark and cool weather were fine; outdoor area underused. Local openers drew 10-15 people combined.', 10, v_keith_p),
  (v_mtg, 'note', 'Bar: Keith reported ~$3,500 spent at the bar with under 100 attendees. Signal''s sales summary (Toast, 11/14) says $3,617.45 collected = $2,860 net sales + $253.35 tax + $480.90 tips; only the $2,860 counts against the $6,000 minimum (48%). That is $29.79 per attendee (96), so $6,000 needs ~200 people and $5,000 ~170 - which is where the 200 target lands. Renegotiate the split or bring an audience that can meet the minimum.', 11, v_keith_p),
  (v_mtg, 'note', 'Bar by hour (Signal): 3-4 PM $95, 4-5 PM $330, 5-6 PM $585, 6-7 PM $343, 7-8 PM $796 (peak), 8-9 PM $476, 9-10 PM $235. The first hour is nearly empty - the case for the brunch hour.', 16, v_keith_p),
  (v_mtg, 'note', 'Paid Instagram/YouTube ads brought bot followers; stop them and build real advocates instead.', 12, v_keith_p),
  (v_mtg, 'note', 'Emmy Adelle agreed to return as resident headliner. Opening back-to-back sets to keep stress low.', 13, v_keith_p),
  (v_mtg, 'note', 'Brunch hour idea: buy food from Crossroads for a 3-6 PM brunch with free sandwiches and drink tickets to pull people in early.', 14, v_keith_p),
  (v_mtg, 'note', 'Mailing list is at 115 subscribers.', 15, v_keith_p);

  for r in
    select * from (values
      -- ev, title, description, owner, priority
      ('di4','Confirm DI #4 and #5 dates with Kim at Signal','Hold Saturday March 6 and Saturday October 9, 2027.','keith','high'),
      ('di4','Propose a 2-event deal to Signal at a $5,000 minimum each','Commit to DI #4 + #5 in advance for a $5,000 minimum spend per event. DI #3 was $6,000; the bar rang $2,860 net (Signal sales summary) and the $2,500 deposit was forfeited. At DI #3''s $29.79 per attendee, $5,000 needs ~170 people.','keith','high'),
      ('di4','Ask Erica about the sponsorship coordinator role','A few hours a week of sponsor outreach. If not her, ask whether she knows someone connected to the cause.','keith','high'),
      ('di4','Email AOM about a $2,500 title sponsorship',null,'keith','medium'),
      ('di4','Talk to NYC Rave Girls (Cassenya) about a post-event party','Invite them to experience Dance Infusion first-hand.','keith','medium'),
      ('di4','Share the previous event survey with the team','Starting point for Martin''s impact-report survey questions.','keith','medium'),
      ('di4','Get a quote from Crossroads for a 3-6 PM brunch hour','Free sandwiches + drink tickets to reward early arrival. Proposed in the meeting, not yet decided.','keith','medium'),
      ('di4','Reach out to university public health departments','Propose student-led capstone projects: research, marketing leads, sponsorship groundwork, impact report. Brooklyn and Queens first.','martin','high'),
      ('di4','Draft impact-report survey questions','Why attendees come and what they get out of it.','martin','medium'),
      ('di4','Research public health schools and MS communities in the boroughs',null,'martin','medium'),
      ('di4','Connect with people at those schools and MS communities',null,'martin','medium'),
      ('di4','Research which channels attendees use to pick events','Feeds the shift away from paid ads.','henry','medium'),
      ('di4','Reach out to local DJs about opening-set collaborations','Cross-promotion: we open their parties (with QR codes), they open ours.','henry','medium'),
      ('di4','Talk to Alex (Signal) about food service','Alex cooks at Signal on Friday nights; possible food partner.','henry','medium'),
      ('di3','Tell the production team the DI #3 audio needs editing for set swaps','Henry put the audio in the Drive raw folder as two long sessions (split at the quiet break between Scott and Keith); set swaps need close listening or runtime maths.','keith','medium'),
      ('di3','Resolve the $500 tip overcharge with Signal (Gavin)','Keith: Gavin charged $500 in tips on top of the $540 already paid; confirmed with Izzy. Signal''s own sales summary shows guests tipped $480.90 at the bar (cards $484.50, 17%), so a $500 tip guarantee was nearly met by guests already. The books record the tip guarantee as $500 final (250) - correct them once this is settled.','keith','high'),
      ('di3','Ask Signal whether the 39 voided drink-ticket items counted toward the minimum','Signal''s summary shows 39 items voided ($509) - the same count as the 39 prepaid drink tickets ($500). Voided items are NOT in the $2,860 net. If those were our drink tickets, the minimum count should be ~$3,369. Ask Gavin/Kim; it matters for the DI #4/#5 deal terms.','keith','medium')
    ) t(ev, title, descr, who, prio)
  loop
    i := i + 1;
    v_ev := case r.ev when 'di4' then v_di4 else v_di3 end;
    insert into tasks (event_id, title, description, status, priority, source, meeting_id, created_by)
    values (v_ev, r.title, r.descr, 'todo', r.prio, 'meeting', v_mtg, v_keith)
    returning id into v_task;
    insert into task_assignments (task_id, actor_id, role)
    values (v_task, case r.who when 'keith' then v_keith when 'henry' then v_henry else v_martin end, 'doer');
    insert into meeting_notes (meeting_id, kind, body, task_id, sort, created_by)
    values (v_mtg, 'action', r.title, v_task, 100 + i, v_keith_p);
  end loop;

  -- asserts
  if (select count(*) from tasks where meeting_id = v_mtg and deleted_at is null) <> 17 then raise exception '256: expected 17 tasks'; end if;
  if (select count(*) from tasks where meeting_id = v_mtg and event_id = v_di4) <> 14 then raise exception '256: expected 14 DI #4 tasks'; end if;
  if (select count(*) from task_assignments ta join tasks t on t.id = ta.task_id where t.meeting_id = v_mtg) <> 17 then raise exception '256: expected 17 assignments'; end if;
  if (select count(*) from meeting_notes where meeting_id = v_mtg) <> 28 then raise exception '256: expected 28 meeting notes'; end if;

  -- DI #3: replace the remembered "~$3,000" with Signal's figure. Money columns untouched.
  update event_control
     set notes = replace(notes, 'bar rang ~$3,000 (Keith, 2026-11-18)',
                 'bar rang $2,860 net per Signal''s sales summary (Toast, 11/14: $3,617.45 collected incl. $253.35 tax + $480.90 tips; 39 items / $509 voided, likely the drink tickets)'),
         updated_at = now()
   where event_id = v_di3;
  if (select notes from event_control where event_id = v_di3) not like '%$2,860 net%' then
    raise exception '256: DI #3 bar note not updated';
  end if;
end $$;

commit;
