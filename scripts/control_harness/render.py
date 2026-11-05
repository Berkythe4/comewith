"""Render control.html headless against a PROD snapshot - see it before you ship it.

    SBP_REF=yaytdosxfhcqatmhctzk python scripts/control_harness/render.py [slug] [section ...]

Defaults: slug dance-infusion-3, sections command door rsvps.

How it works: snapshot.sql reads the event's rows (one Management API call,
read-only), the page is copied into _build/ with its Supabase client swapped for
fake-supabase.js (answers selects from the snapshot, ignores writes), served on
localhost, and screenshotted with headless Edge. PNGs land in _build/.

_build/ holds GUEST DATA (names, emails) and is gitignored. Never commit it.
Added 2026-11-05 (LEARNINGS 106): CLAUDE.md said there was no local way to see
this page; there is.
"""
import json, os, re, shutil, subprocess, sys, threading, time, http.server, functools
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HERE = Path(__file__).resolve().parent
BUILD = HERE / '_build'
EDGE = [r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', r'C:\Program Files\Microsoft\Edge\Application\msedge.exe']

def main():
    args = sys.argv[1:]
    slug = args[0] if args and not args[0] in ('command', 'door', 'rsvps', 'sponsors', 'tickets', 'drinks', 'finance', 'audience', 'exec') else 'dance-infusion-3'
    secs = [a for a in args if a != slug] or ['command', 'door', 'rsvps']
    if not re.fullmatch(r'[a-z0-9-]+', slug):
        sys.exit('slug must be lowercase letters, digits and dashes')
    BUILD.mkdir(exist_ok=True)
    (BUILD / 'assets').mkdir(exist_ok=True)
    sql = (HERE / 'snapshot.sql').read_text(encoding='utf-8').replace(':SLUG', f"'{slug}'")
    (BUILD / 'snapshot.sql').write_text(sql, encoding='utf-8')
    out = subprocess.run([sys.executable, str(ROOT / 'db.py'), str(BUILD / 'snapshot.sql')], capture_output=True, text=True, cwd=ROOT)
    if out.returncode:
        sys.exit(out.stderr)
    snap = {k: (v or []) for k, v in json.loads(out.stdout)[0]['r'].items()}
    (BUILD / 'snap.json').write_text(json.dumps(snap), encoding='utf-8')
    for f in ('control-center.js', 'import-engine.js'):
        shutil.copy(ROOT / 'assets' / f, BUILD / 'assets' / f)
    shutil.copy(HERE / 'fake-supabase.js', BUILD / 'fake-supabase.js')
    page = (ROOT / 'control.html').read_text(encoding='utf-8')
    page = page.replace('https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm', './fake-supabase.js')
    (BUILD / 'control.html').write_text(page, encoding='utf-8')

    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(BUILD))
    handler.log_message = lambda *a: None
    srv = http.server.ThreadingHTTPServer(('127.0.0.1', 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    port = srv.server_address[1]
    edge = next((e for e in EDGE if os.path.exists(e)), None)
    if not edge:
        sys.exit('Edge not found - open http://127.0.0.1:%d/control.html?event=%s yourself' % (port, slug))
    try:
        for sec in secs:
            png = BUILD / f'{slug}_{sec}.png'
            subprocess.run([edge, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--window-size=1280,2600',
                            '--virtual-time-budget=8000', f'--screenshot={png}', f'http://127.0.0.1:{port}/control.html?event={slug}#{sec}'],
                           capture_output=True, timeout=90)
            print('wrote', png)
        # PHONE: headless Edge will not lay out narrower than ~477px whatever
        # --window-size says (it crops the screenshot instead, which LOOKS like
        # overflow). A 390px iframe gives a true phone viewport.
        (BUILD / 'phone.html').write_text('<!doctype html><meta charset="utf-8"><body style="margin:0;background:#000">'
            f'<iframe style="width:390px;height:2400px;border:0" src="control.html?event={slug}#{secs[0]}"></iframe>', encoding='utf-8')
        mobile = BUILD / f'{slug}_{secs[0]}_phone.png'
        subprocess.run([edge, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--window-size=600,2400',
                        '--virtual-time-budget=10000', f'--screenshot={mobile}', f'http://127.0.0.1:{port}/phone.html'],
                       capture_output=True, timeout=90)
        print('wrote', mobile, '(the phone is the left 390px)')
    finally:
        srv.shutdown()

if __name__ == '__main__':
    main()
