-- 213: repair campaign stats broken by Resend's own "email.sent" webhook.
--
-- Once the webhook began forwarding email.sent, resend-webhook stored it as a SECOND
-- 'sent' row beside the one send-campaign writes. Two effects:
--   1. Campaign stats counted 'sent' rows raw, so every campaign showed 2x sent.
--   2. The webhook's attribution lookup used .maybeSingle() on 'sent', which errors on
--      two rows -> campaign_id/subscriber_id null on every later delivered/opened/
--      clicked event. DI#3 Save the Date showed 11 delivered of 87 (really 87).
-- The function now ignores email.sent (deployed alongside). This re-attributes the
-- orphans and removes the duplicate 'sent' rows. Data only; no schema change.
--
-- Our own 'sent' rows are told apart by metadata: send-campaign leaves it '{}' (or
-- {email} for a CC), the webhook stores Resend's payload, which always has created_at.

begin;

with ours as (
  select resend_event_id, campaign_id, subscriber_id
  from public.mailing_events
  where event_type = 'sent' and campaign_id is not null
    and resend_event_id is not null and not (metadata ? 'created_at')
)
update public.mailing_events e
set campaign_id = o.campaign_id, subscriber_id = o.subscriber_id
from ours o
where e.resend_event_id = o.resend_event_id
  and e.campaign_id is null
  and e.event_type <> 'sent';

-- Only duplicates that shadow one of our campaign sends. Webhook 'sent' rows for
-- non-campaign mail (actor emails, invoices) carry no campaign and are left alone.
delete from public.mailing_events w
where w.event_type = 'sent' and w.metadata ? 'created_at'
  and exists (
    select 1 from public.mailing_events o
    where o.resend_event_id = w.resend_event_id and o.event_type = 'sent'
      and o.campaign_id is not null and not (o.metadata ? 'created_at')
  );

do $$
declare n int;
begin
  select count(*) into n from public.mailing_events w
  where w.event_type = 'sent' and w.campaign_id is not null and w.metadata ? 'created_at';
  if n > 0 then raise exception '213: % webhook sent rows still attributed to a campaign', n; end if;
end $$;

commit;
