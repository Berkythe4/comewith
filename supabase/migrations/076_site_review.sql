-- =============================================================================
-- 076_site_review.sql
-- Site Review log: findings from the 2026-07-08 full-site audit, kept as live,
-- editable rows so review/maintenance work happens inside the dashboard.
-- New module 'site-review' sits directly under Site Editor (Insights group).
-- Also seeds site_content 'ops.ra_guestlist_type' so the RA guest-list export
-- "Type" column is editable in Site Editor → Dashboard settings.
-- Grants: inherited from 013 default privileges — no explicit grants here.
-- =============================================================================
begin;

create table if not exists public.site_review_items (
  id uuid primary key default gen_random_uuid(),
  kind text not null check (kind in ('bug','improvement','saved','capability','data')),
  area text not null,
  title text not null,
  detail text,
  file_ref text,
  status text not null default 'open' check (status in ('fixed','open','review','planned','dismissed')),
  sort integer not null default 100,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.site_review_items enable row level security;
drop policy if exists "Admins manage site review" on public.site_review_items;
create policy "Admins manage site review" on public.site_review_items
  for all using (public.is_admin()) with check (public.is_admin());

-- Nav entry: right below Site Editor (193) in Insights. signed_off=false keeps
-- it master-only until released from Team → Modules.
insert into public.module_registry (key, label, nav_group, sort_order, built, signed_off, master_only, default_roles)
values ('site-review', 'Site Review', 'Insights', 194, true, false, false, '{marketing,full}')
on conflict (key) do nothing;

-- RA export Type setting (read by the dashboard's Export-for-RA; editable in Site Editor).
insert into public.site_content (key, value)
values ('ops.ra_guestlist_type', 'Guest List')
on conflict (key) do nothing;

-- ---------------------------------------------------------------------------
-- Seed: findings from the 2026-07-08 audit (5 parallel reviews + DB checks).
-- ---------------------------------------------------------------------------
insert into public.site_review_items (kind, area, title, detail, file_ref, status, sort) values
-- FIXED in this sweep
('bug','Inquiries','Convert-to-client could half-complete and duplicate people','actor_roles/inquiries writes were unchecked (a failure could leave an actor without the customer role or the inquiry unconverted), and converting an inquiry whose email already belonged to an actor created a twin. Now: dedupe-by-email offers to link the existing actor, and every write is error-checked.','dashboard.html convertInquiry()','fixed',10),
('bug','Event hub','Equipment load-in checkbox stayed checked when the save failed','The box now reverts if the database write errors, so the screen never claims gear is loaded when it is not.','dashboard.html hubToggleLoaded()','fixed',20),
('bug','Agreements','send-agreement did not check the status update after emailing','If flipping the agreement to ''sent'' failed after a successful email, the dashboard showed a stale status with no warning. The function now reports it (email out + status warning). Redeployed.','supabase/functions/send-agreement','fixed',30),
('improvement','Event hub','Fee-to-expense error now says what to do','Was "No fee on this participant"; now explains: edit the participant, add the fee, then post it.','dashboard.html hubFeeToExpense()','fixed',40),
('improvement','Public site','Social-share previews (og:image / og:url / twitter:card) added','index.html, watch.html and artist.html had no og:image, so links shared to IG/WhatsApp/Discord/X rendered bare. All three now carry the brand logo + canonical URL.','index.html, watch.html, artist.html <head>','fixed',50),
('improvement','Customer portal','Empty state no longer dead-ends','"No agreements yet" now offers mailto:berky@comewith.org.','customer_portal.html','fixed',60),
('capability','Guest list','RA export "Type" editable in-app','The Resident Advisor CSV export reads its Type column from Site Editor → Dashboard settings (ops.ra_guestlist_type) instead of a hardcoded ''Guest List''.','Site Editor → Dashboard settings','fixed',70),
-- OPEN data hygiene (your call)
('data','Actors','Duplicate email: "Victoriarose Vargas" and "Miss Vee"','Both actors carry victoriarose.business@outlook.com. If they are the same person, merge (keep the richer record, repoint roles/participants); if not, fix one email.','Actors tab','open',100),
('data','Events','3 events have no venue set','Henry Artist Showcase (5/17), Knicks G5 Watch Party (6/13), July 4th Weekend (7/4). Backfill from the event hub → Set venue if you want venue history/capacity autofill to work for them.','Events → each event → Overview','open',110),
('data','Public site','og:image is pinned to the current logo file','Static pages cannot read the CMS, so the share image points at the logo''s current storage URL. If you re-upload a logo, the URL changes and this needs a one-line update in the three page heads.','index/watch/artist.html','open',120),
-- SAVED for review (bigger or judgment calls)
('saved','Security','No rate limiting on public endpoints (subscribe, inquiry-notify)','A bot could spam confirmation emails or admin notifications. Fix idea: per-email/IP throttle (e.g. max 3 sends per address per hour) checked against recent rows before sending.','supabase/functions/subscribe, inquiry-notify','review',200),
('saved','Edge functions','Error messages can leak internal detail','Several functions return raw DB/Resend error text to the client. Fine while you are the only user; sanitize (generic message out, detail to logs) before external users.','supabase/functions/*','review',210),
('saved','Edge functions','FROM address + SITE_URL duplicated across ~8 functions','"Come With <berky@comewith.org>" and the site URL are pasted per-function. Centralize into secrets (FROM_EMAIL) so a future address change is one edit.','supabase/functions/*','review',220),
('saved','Email','Outbound email copy is hardcoded','Subjects/bodies for the artist update link, intake invite, social-calendar snapshot and agreement emails live in code. An in-app "Email templates" editor (like Templates for outreach) would make them yours to edit.','dashboard.html + edge functions','review',230),
('saved','Vendors','Vendor category list is hardcoded','VENDOR_CATEGORIES (10 entries) lives in code; make it editable if your categories evolve.','dashboard.html VENDOR_CATEGORIES','review',240),
('saved','Social calendar','Post "series" tag only offers Parties / Dance Infusion','Production & Content Creation posts currently tag as "general". Add the two series if you start planning content for them.','dashboard.html openPostModal()','review',250),
('saved','KPIs','KPI views cover Parties + DI only (by design)','Production/content events roll up on the Events page money models instead. Revisit only if those series need Strategy cards.','v_kpi_* views','review',260),
('saved','Homepage','Offline fallback arrays are hardcoded','PAST recaps + DJS fallbacks in index.html only render if Supabase is unreachable; harmless, but they will age. CMS-drive or trim them someday.','index.html PAST/DJS consts','review',270),
('saved','Impact report','Old archive copies contain outdated hardcoded numbers','events/dance-infusion/di-02-2026-05/** archives predate the Supabase-driven live report. Not served as the live report; consider pruning archives from the repo to avoid confusion.','events/dance-infusion/di-02-2026-05/','review',280);

commit;
-- POST: table + policy live; module 'site-review' appears (master-only until
-- signed off); Site Editor gains "Dashboard settings"; 19 findings seeded.
