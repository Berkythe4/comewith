#!/usr/bin/env python3
"""
One-shot: apply migration 067, seed the DI#2 impact_report jsonb, and set the
internal KPI target to 50%-to-mission. Idempotent.

Run from repo root:  python scripts/apply_067_impact_report.py

Needs a VALID SBP_PAT in .env (Supabase Management API token for prod).
If the token is expired you'll get HTTP 401 — refresh it at
https://supabase.com/dashboard/account/tokens and paste into .env (SBP_PAT=...).
"""
import json, os, sys, urllib.request, urllib.error

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

def env():
    e = {}
    for line in open(os.path.join(ROOT, '.env'), encoding='utf-8'):
        line = line.strip()
        if '=' in line and not line.startswith('#'):
            k, v = line.split('=', 1); e[k] = v.strip()
    return e

E = env()
REF, PAT = E['SBP_REF_PROD'], E['SBP_PAT']

def q(sql):
    req = urllib.request.Request(
        'https://api.supabase.com/v1/projects/%s/database/query' % REF,
        data=json.dumps({'query': sql}).encode(),
        headers={'Authorization': 'Bearer ' + PAT, 'Content-Type': 'application/json', 'User-Agent': UA},
        method='POST')
    try:
        return json.load(urllib.request.urlopen(req))
    except urllib.error.HTTPError as ex:
        body = ex.read().decode(errors='replace')
        print('  HTTP %s: %s' % (ex.code, body[:300]))
        if ex.code == 401:
            print('\n>>> SBP_PAT is expired/invalid. Refresh it and re-run. <<<')
        sys.exit(1)

def jload(p):
    return json.load(open(os.path.join(ROOT, p), encoding='utf-8'))

# ---- build the seed jsonb from the local source-of-truth JSON + copy changes ----
def build_impact_report():
    d2 = jload('events/dance-infusion/di-02-2026-05/data/dance_infusion.json')
    c  = jload('events/dance-infusion/di-02-2026-05/data/impact_report_content.json')
    d1 = jload('events/dance-infusion/di-01-2025-09/data/dance_infusion_di1.json')
    # DI#1 copy decisions (Keith): attendance shown as 42, sponsors = 0.
    d1.setdefault('attendance', {})['attendance'] = 42
    d1['sponsors'] = {'count': 0, 'note': 'Dance Infusion #1 was solo-run with no sponsors.'}
    nar = c.get('narratives', {})
    content = {
        'event': c.get('event', {}),
        'narratives': {  # reach removed by design
            'why_we_dance': nar.get('why_we_dance', ''),
            'the_event': nar.get('the_event', ''),
            'whats_next': nar.get('whats_next', ''),
        },
        'human_moment': {'quote': '', 'attribution': ''},  # blank until consent
        'sponsor_display': c.get('sponsor_display', []),
        'team': c.get('team', []),
        'artists': c.get('artists', []),
        'contact': c.get('contact', {}),
        'photos': {'hero': None, 'inline': None},
    }
    return {
        'config': d2.get('config', {}),          # incl. attendees_confirmed=117
        'sponsors': d2.get('sponsors', []),      # count -> 9
        'dj_allocations': d2.get('dj_allocations', []),
        'expenses': d2.get('expenses', []),      # e1..e6 -> audit money flow
        'content': content,
        'di1': d1,
        'audit': {'target_pct': 50},             # forward goal = 50% to mission
    }

def sql_str(s):
    return "'" + s.replace("'", "''") + "'"

def main():
    print('1) Applying migration 067 ...')
    q(open(os.path.join(ROOT, 'supabase/migrations/067_impact_report.sql'), encoding='utf-8').read())
    print('   columns + view created.')

    print('2) Locating the DI#2 event ...')
    rows = q("select id, name from public.events where series='Dance Infusion' and deleted_at is null "
             "order by (name ilike '%#2%') desc, event_date desc limit 1")
    # Management API returns a list of result rows.
    ev = rows[0] if isinstance(rows, list) and rows else None
    if not ev:
        print('   No Dance Infusion event found — aborting seed.'); sys.exit(1)
    eid, ename = ev['id'], ev['name']
    print('   -> %s (%s)' % (ename, eid))

    print('3) Seeding events.impact_report (impact_report_public left FALSE) ...')
    payload = json.dumps(build_impact_report())
    q("update public.events set impact_report = %s::jsonb where id = '%s'" % (sql_str(payload), eid))
    print('   seeded %d bytes.' % len(payload))

    print('4) Setting internal KPI target di.cost_to_raise = 0.50 (=50%% to mission) ...')
    q("insert into public.kpi_targets (metric_key, workstream, label, target_value, comparison, unit, effective_date) "
      "select 'di.cost_to_raise','dance_infusion','Cost to raise $1',0.50,'lte','$',current_date "
      "where not exists (select 1 from public.kpi_targets where metric_key='di.cost_to_raise' "
      "and target_value=0.50 and effective_date=current_date)")
    print('   target set.')

    print('5) Verifying ...')
    chk = q("select count(*) as n from information_schema.columns where table_schema='public' "
            "and table_name='events' and column_name in ('impact_report','impact_report_public')")
    print('   event columns present:', chk)
    pub = q("select impact_report_public from public.events where id='%s'" % eid)
    print('   impact_report_public (expect false):', pub)
    print('\nDONE. To go live: flip "Show on public site" in the dashboard event hub')
    print('(or run: update public.events set impact_report_public=true where id=\'%s\';)' % eid)

if __name__ == '__main__':
    main()
