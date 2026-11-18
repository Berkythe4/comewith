-- 255: DI #3 - impact report draft, Keith's closing sponsorship, event closed (Keith, 2026-11-18).
--
-- 1. events.impact_report for DI #3 (rendered by
--    events/dance-infusion/di-03-2026-11/reports/impact-report.html). impact_report_public
--    stays FALSE: publishing is Keith's toggle in the dashboard, never a side effect.
--    Figures: to MS $1,965 = $25 x 65 tickets + $340 donations (the $40 tier keeps
--    $15 for production - Keith kept the split 2026-11-18); 96 on the floor (door list);
--    three-night total $6,105 = $1,140 + $3,000 + $1,965 (matches homepage di.goal_raised).
--    Narrative copy is a DRAFT for Keith to edit in the dashboard.
-- 2. Keith will send $896.95 to cover what is left (Frick Frack $500 + MS $1,965 minus
--    DI Bluevine after the $770 Zeffy payout). Booked as a CONFIRMED (committed, not yet
--    received) Founder sponsorship; mark it paid when the transfer lands.
-- 3. DI #3 is over: events.status -> completed; homepage "Get tickets" (di.ticket_show,
--    pointing at the DI #3 Zeffy page) -> off.

begin;

do $$ begin
  if (select impact_report from events where id = 'f13bf700-7383-497b-91b8-d1f6781e077b') <> '{}'::jsonb
  then raise exception 'DI #3 already has an impact report'; end if;
  if exists (select 1 from sponsorships where event_id = 'f13bf700-7383-497b-91b8-d1f6781e077b' and cash_amount = 896.95)
  then raise exception 'closing sponsorship already booked'; end if;
end $$;

update events
   set impact_report = $json${
  "config": {
    "event_name": "Dance Infusion #3",
    "to_mission": 1965,
    "to_mission_formula": "$25 x 65 paid tickets ($1,625) + $340 Zeffy donations. The $40 tier sends $25 to MS and keeps $15 for production (Keith kept the split 2026-11-18).",
    "attendance": 96,
    "attendance_note": "Door list: 64 paid + 32 comp. Zeffy check-in scanned only 44 of 62 tickets, so it is not a headcount.",
    "tickets_sold": 65,
    "artists_count": 5,
    "cumulative_to_mission": 6105,
    "growth_lede": "Every Dance Infusion so far, side by side.",
    "growth": [
      { "label": "To the MS Society", "kind": "money", "values": { "#1": 1140, "#2": 3000, "#3": 1965 } },
      { "label": "On the floor", "kind": "count", "values": { "#1": 42, "#2": 117, "#3": 96 } }
    ],
    "internal_status": "DRAFT 2026-11-18. MS $1,965 not yet sent; Frick Frack $500 unpaid; Keith's $896.95 closing sponsorship committed, not received."
  },
  "content": {
    "event": {
      "name": "Dance Infusion #3",
      "date": "November 14, 2026",
      "venue": "Signal NYC, Brooklyn",
      "tagline": "Dancing for a Cure, One Beat at a Time",
      "hero_subtitle": "A third afternoon on the floor for the National MS Society.",
      "night_heading": "November 14 at Signal NYC"
    },
    "narratives": {
      "why_we_dance": "Dance Infusion exists because MS doesn't stop, and neither do we. The National MS Society funds research, advocacy and support for the nearly one million Americans living with multiple sclerosis. Every ticket to Dance Infusion sends $25 to the Society, and every donation goes to it in full.",
      "the_event": "On November 14, Dance Infusion came back to Signal NYC for its third night. From 3PM to 10PM, five DJs (Miss Vee, Rainbow TuTu, Berky, Emmy Adelle and KRNeY) kept the room moving, and 96 people came through the door. Four Champion sponsors and our donors stood behind it.",
      "whats_next": "Dance Infusion will be back. Follow @DanceInfusionNYC for the next date, sponsorship slots and early tickets. If you danced with us in November, thank you."
    },
    "human_moment": { "quote": "", "attribution": "" },
    "photos": {},
    "artists": ["Miss Vee", "Rainbow TuTu", "Berky", "Emmy Adelle", "KRNeY"],
    "team": [
      { "name": "Keith Berkman (Berky)", "role": "Producer" },
      { "name": "Henry Zaradich", "role": "Producer" },
      { "name": "Janelle Sochet", "role": "Marketing Director" }
    ],
    "sponsor_display": [
      { "tier": "Champions", "members": [
        { "name": "Pulse Devices" }, { "name": "Theresa Berkman" },
        { "name": "Patrick Savery & Dr. Arnaldo Martinez" }, { "name": "Keith Berkman" } ] },
      { "tier": "Founding sponsors", "members": [ { "name": "Come With" }, { "name": "Keith Berkman" } ] }
    ],
    "contact": { "email": "danceinfusionnyc@gmail.com", "instagram": "@DanceInfusionNYC" }
  }
}$json$::jsonb
 where id = 'f13bf700-7383-497b-91b8-d1f6781e077b';

insert into sponsorships (event_id, actor_id, tier, cash_amount, status, notes, payment_refs)
values ('f13bf700-7383-497b-91b8-d1f6781e077b', 'ca4a6e21-de18-45d5-a4fd-1d370b6603bb', 'Founder sponsorship', 896.95, 'confirmed',
        'Keith''s closing founder sponsorship, committed 2026-11-18, NOT yet received: covers what DI #3 still owes (Frick Frack $500 + MS Society $1,965 = $2,465) minus DI Bluevine after the pending $770 Zeffy payout ($798.05 + $770 = $1,568.05). Mark paid when the transfer lands. Not a donation.',
        '{}');

update events set status = 'completed' where id = 'f13bf700-7383-497b-91b8-d1f6781e077b';

update site_content set value = '0' where key = 'di.ticket_show';   -- text column; same form as di.donate_show

do $$ begin
  if (select impact_report_public from events where id = 'f13bf700-7383-497b-91b8-d1f6781e077b')
  then raise exception 'DI #3 report must stay unpublished'; end if;
  if (select (impact_report->'config'->>'to_mission')::numeric from events where id = 'f13bf700-7383-497b-91b8-d1f6781e077b') <> 1965
  then raise exception 'to_mission wrong'; end if;
end $$;

commit;
