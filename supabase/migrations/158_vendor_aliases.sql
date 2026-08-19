-- ============================================================
-- COME WITH — 158 vendor aliases, and non-event F&B as networking
--
-- THE PROBLEM. 263 expenses carry 72 distinct payee strings, and the same vendor
-- arrives spelled several ways depending on which feed saw it:
--   Ableton            | Ableton Ag | Ableton Ag Deu
--   Anthropic Ca       | Anthropic* Claude Sub | Claude.ai Subscription Ca
--   Bandcamp           | Bandcamp Ventures LLC | Bandcampeffy | Bandcampsinderlyn …
--   Beatport, Google ADS…, GOOGLE *Workspace_come, Splice.com Ny, Sq *cross X Roads
-- 34 canonical vendor actors ALREADY EXIST; the variants simply were not linked,
-- leaving 131 of 263 expenses with no vendor_actor_id.
--
-- THE FIX IS A RULE, NOT A CLEANUP. A one-off UPDATE would fix today and break
-- again on the next import, because bank feeds keep inventing new spellings.
-- vendor_aliases maps a lowercase substring to an actor; resolve_vendor_actor()
-- applies it, longest pattern first so 'claude.ai' beats a broader rule. Backfill
-- runs it over history, and the importer can call the same function forever.
--
-- Also: non-event food and beverage becomes 'Marketing / Networking'. Coffees and
-- festival bar tabs are relationship spend, not operations - Keith's call. Rows
-- attached to an event keep their event category; those really are cost of the
-- night.
-- ============================================================
begin;

-- ---------------------------------------------------------------
-- 1. The alias table
-- ---------------------------------------------------------------
create table if not exists public.vendor_aliases (
  id         uuid primary key default gen_random_uuid(),
  pattern    text not null unique,          -- lowercase substring, matched with LIKE
  actor_id   uuid not null references public.actors(id) on delete cascade,
  note       text,
  created_at timestamptz not null default now()
);

create index if not exists idx_vendor_aliases_actor on public.vendor_aliases(actor_id);

alter table public.vendor_aliases enable row level security;
drop policy if exists "Admins manage vendor aliases" on public.vendor_aliases;
create policy "Admins manage vendor aliases" on public.vendor_aliases
  for all using (public.is_admin());
revoke all on public.vendor_aliases from anon;

comment on table public.vendor_aliases is
  'Payee spelling -> canonical actor. Bank and card feeds render the same vendor '
  'differently every time; this is the rule that survives the next import.';

-- ---------------------------------------------------------------
-- 2. Resolver — longest pattern wins
-- ---------------------------------------------------------------
create or replace function public.resolve_vendor_actor(p_vendor text)
returns uuid language sql stable as $$
  select a.actor_id
    from public.vendor_aliases a
   where p_vendor is not null
     and lower(p_vendor) like '%' || a.pattern || '%'
   order by length(a.pattern) desc
   limit 1;
$$;

comment on function public.resolve_vendor_actor(text) is
  'Longest matching alias wins, so a specific rule (claude.ai) beats a general '
  'one. Returns null when nothing matches rather than guessing.';

-- ---------------------------------------------------------------
-- 3. Seed the aliases against the actors that already exist
-- ---------------------------------------------------------------
-- Any canonical vendor with no actor yet is created first, so every pattern has
-- somewhere to point.
insert into public.actors (display_name, status)
select v.name, 'active'
  from (values
    ('Best Buy'), ('Serato'), ('RepostExchange'), ('Green Room NYC'),
    ('Park Slope Convenience'), ('Support Women DJs'), ('Venmo'), ('Facebook')
  ) as v(name)
 where not exists (
   select 1 from public.actors a where a.deleted_at is null and lower(a.display_name) = lower(v.name));

insert into public.vendor_aliases (pattern, actor_id, note)
select v.pattern, a.id, 'seeded by 158'
  from (values
    -- software / subscriptions
    ('ableton',            'Ableton'),
    ('anthropic',          'Anthropic'),
    ('claude.ai',          'Anthropic'),
    ('beatport',           'Beatport'),
    ('bandcamp',           'Bandcamp'),
    ('splice',             'Splice'),
    ('soundcloud',         'SoundCloud'),
    ('netlify',            'Netlify'),
    ('namecheap',          'Namecheap'),
    ('serato',             'Serato'),
    ('rekordbox',          'rekordbox / Pioneer DJ'),
    ('pioneer',            'rekordbox / Pioneer DJ'),
    ('repostexchan',       'RepostExchange'),
    ('google',             'Google'),
    -- marketing / platforms
    ('resident advisor',   'Resident Advisor'),
    ('meta / instagram',   'Meta'),
    ('facebook',           'Facebook'),
    ('support women djs',  'Support Women DJs'),
    -- gear
    ('sweetwater',         'Sweetwater (Benjamin Denen)'),
    ('b&h photo',          'B&H Photo'),
    ('guitar center',      'Guitar Center'),
    ('best buy',           'Best Buy'),
    ('krk',                'KRK Systems'),
    ('temu',               'Temu'),
    ('amazon music',       'Amazon'),
    ('amazon prime',       'Amazon (Prime Video)'),
    ('prime video',        'Amazon (Prime Video)'),
    ('amazon',             'Amazon'),
    -- venues / places
    ('signal',             'Signal NYC'),
    ('refuge',             'Refuge Nightclub'),
    ('acoustik garden',    'Acoustik Garden Lounge'),
    ('kingdomflush',       'Kingdomflush LLC'),
    ('tce presents',       'TCE Presents'),
    ('zion',               'Zion'),
    ('cross x roads',      'Crossroads Café'),
    ('crossroads',         'Crossroads Café'),
    ('green room',         'Green Room NYC'),
    ('park slope convenience', 'Park Slope Convenience'),
    ('333 stagg',          '333 Stagg (Afterparty)'),
    ('elements',           'Elements Music & Arts Festival'),
    -- people
    ('janelle',            'Janelle Sochet'),
    ('sochetjanel',        'Janelle Sochet'),
    ('henry',              'Henry'),
    ('19th & 7th',         '19th & 7th Productions (Michael McManus)'),
    ('wellness pharmacy',  'Wellness Pharmacy'),
    ('uber',               'Uber'),
    ('venmo',              'Venmo')
  ) as v(pattern, actor_name)
  join public.actors a
    on a.deleted_at is null and lower(a.display_name) = lower(v.actor_name)
on conflict (pattern) do nothing;

-- ---------------------------------------------------------------
-- 4. Backfill every unlinked expense
-- ---------------------------------------------------------------
update public.expenses e
   set vendor_actor_id = public.resolve_vendor_actor(e.vendor)
 where e.deleted_at is null
   and e.vendor_actor_id is null
   and public.resolve_vendor_actor(e.vendor) is not null;

-- ---------------------------------------------------------------
-- 5. Non-event food & beverage -> Marketing / Networking
-- ---------------------------------------------------------------
-- Coffees, bar tabs and convenience runs that belong to no event are relationship
-- spend. Event-linked F&B is left alone: that is genuinely the cost of the night.
update public.expenses
   set category = 'Marketing / Networking'
 where deleted_at is null
   and event_id is null
   and (vendor ilike '%cross x roads%' or vendor ilike '%crossroads%'
     or vendor ilike '%green room%'    or vendor ilike '%elements bars%'
     or vendor ilike '%cafe%'          or vendor ilike '%caf_%'
     or vendor ilike '%grocery%'       or vendor ilike '%food & beverage%'
     or vendor ilike '%convenience%')
   and coalesce(category, '') <> 'Marketing / Networking';

commit;

-- DOWN:
--   drop function if exists public.resolve_vendor_actor(text);
--   drop table if exists public.vendor_aliases;
--   (category and vendor_actor_id changes are data, not schema — restore from backup)
