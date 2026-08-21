-- ============================================================
-- COME WITH — 193 the covering email is worth keeping
--
-- The Send screen collected a cc, a subject and a message and then threw all
-- three away the moment you closed it. So the only way to write a covering note
-- was to write it and send in one sitting, and the only way to cc the bookkeeper
-- was to remember the address every single time.
--
-- TWO LEVELS, because they answer two different questions:
--
--   invoice_settings.default_*   what EVERY invoice should start with
--                                (the bookkeeper's cc, house wording)
--   invoices.send_*              what THIS invoice says
--                                (drafted now, sent later, or re-sent)
--
-- The Send screen fills from the invoice if it has anything, and from the
-- defaults if it does not. Saving on the Send screen writes the invoice row, so
-- "save and go back" keeps a half-written note without issuing anything.
--
-- Sending writes them too, so what went out is on the record rather than being
-- reconstructed from the Resend dashboard.
--
-- NOT a template language. `default_subject` and `default_message` are plain
-- text; the only substitution is the one the code already does when the subject
-- is blank ("Invoice CW-2026-0001 from Come With — $1,414.73 due"). Inventing
-- {{placeholders}} here would be a parser, an escaping problem and a support
-- burden for a field that gets typed once a year.
-- ============================================================
begin;

alter table public.invoices
  add column if not exists send_cc      text,
  add column if not exists send_subject text,
  add column if not exists send_note    text;

alter table public.invoice_settings
  add column if not exists default_cc      text,
  add column if not exists default_subject text,
  add column if not exists default_message text;

commit;

-- DOWN:
--   alter table public.invoices drop column if exists send_cc, drop column if exists send_subject,
--                              drop column if exists send_note;
--   alter table public.invoice_settings drop column if exists default_cc,
--     drop column if exists default_subject, drop column if exists default_message;
