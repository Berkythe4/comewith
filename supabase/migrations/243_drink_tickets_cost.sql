-- =============================================================================
-- 243_drink_tickets_cost.sql - Keith, 2026-11-11 (DI #3)
-- =============================================================================
--   Keith paid $500 for the 38 drink tickets locked with Signal. A per-ticket
--   value can't hold that: numeric(10,2) gives $13.00 (= $494) or $13.16
--   (= $500.08). So the TOTAL is what gets stored, and the per-ticket value is
--   derived from it (total / locked) in assets/control-center.js.
--     * event_control.drink_tickets_cost: total paid for the locked tickets.
--       Null = not set; drink_ticket_value is then used as before.
--     * DI #3: $500; drink_ticket_value kept in step at 13.16 (rounded, for
--       anything reading the raw column; the page uses the exact 500/38).
--   Additive. Table is is_admin() + anon-revoked since 222; nothing to grant.
-- =============================================================================

begin;

alter table public.event_control
  add column drink_tickets_cost numeric(10,2) check (drink_tickets_cost is null or drink_tickets_cost >= 0);

comment on column public.event_control.drink_tickets_cost is
  'Total $ paid for the drink tickets locked with the venue. When set, $ per ticket = this / drink_tickets_locked (drink_ticket_value is then only a rounded copy).';

do $$
declare n int;
begin
  update public.event_control c set drink_tickets_cost = 500, drink_ticket_value = 13.16
    from public.events e
   where e.id = c.event_id and e.slug = 'dance-infusion-3' and c.drink_tickets_locked = 38;
  get diagnostics n = row_count; if n <> 1 then raise exception 'DI #3 control row not as expected (locked <> 38?)'; end if;
end $$;

commit;
