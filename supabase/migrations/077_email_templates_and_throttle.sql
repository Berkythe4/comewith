-- =============================================================================
-- 077_email_templates_and_throttle.sql
-- 1) subscribers.confirm_sent_at — supports the subscribe rate limit (no
--    re-sending a confirm email to the same address within 10 minutes).
-- 2) email_templates — outbound email copy becomes owner-editable in-app
--    (Templates screen). Senders read the row by key and fall back to their
--    current hardcoded copy if the row is missing, so this is zero-risk.
--    {{placeholders}} are substituted at send time; link-ish placeholders
--    render as styled buttons.
-- 3) site_content 'ops.vendor_categories' — vendor category list editable in
--    Site Editor → Dashboard settings.
-- Grants: inherited from 013 default privileges — no explicit grants.
-- =============================================================================
begin;

alter table public.subscribers add column if not exists confirm_sent_at timestamptz;

create table if not exists public.email_templates (
  key text primary key,
  label text not null,
  subject text not null,
  body text not null,
  placeholders text,               -- documentation shown in the editor
  updated_at timestamptz not null default now()
);

alter table public.email_templates enable row level security;
drop policy if exists "Admins manage email templates" on public.email_templates;
create policy "Admins manage email templates" on public.email_templates
  for all using (public.is_admin()) with check (public.is_admin());

insert into public.email_templates (key, label, subject, body, placeholders) values
('artist_update_link', 'Artist — self-update link', 'Update your Come With artist profile',
 E'Hi {{first_name}},\n\nYou can update your Come With artist profile — bio, socials and photo — using your private link below. No login needed.\n\n{{link}}\n\nWhatever you save shows up on your profile page at comewith.org. Thanks!\n— Come With',
 '{{first_name}} = their first name · {{link}} = the private update link'),
('artist_intake_invite', 'Artist — intake form invite', 'Join the Come With collective',
 E'Hi,\n\nWe''d love to add you to the Come With collective. Fill out this quick intake form — bio, socials and a photo — and you''ll be set up. No login needed.\n\n{{link}}\n\nThanks!\n— Come With',
 '{{link}} = the public intake form link'),
('subscribe_confirm', 'Mailing list — confirm subscription', 'Confirm your Come With subscription',
 E'{{greeting}}\n\nYou signed up for the Come With mailing list. One click and you''re in:\n\n{{confirm_button}}\n\nIf you didn''t sign up, ignore this email and you won''t be subscribed.',
 '{{greeting}} = "Hi <name>," or "Hi," · {{confirm_button}} = the Confirm subscription button'),
('survey_invite', 'Survey — personal invite email', '{{survey_title}}',
 E'Hi {{name}},\n\n{{intro}}\n\n{{button}}',
 '{{name}} = recipient name · {{intro}} = the survey''s intro text · {{button}} = the Take-the-survey button · subject {{survey_title}} = the survey title')
on conflict (key) do nothing;

insert into public.site_content (key, value)
values ('ops.vendor_categories', 'Software & Subscriptions · Equipment & Gear · Marketing & Ads · Travel · Talent & Contractors · Venue & Space · Supplies · Professional Development · Platform & Fees · Other')
on conflict (key) do nothing;

commit;
-- POST: templates editable on the Templates screen; subscribe throttle active
-- after the function redeploy; vendor categories editable in Site Editor.
