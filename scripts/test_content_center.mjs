// Runs the Event > Content control-center builders against fixtures.
//
//   node scripts/test_content_center.mjs        (from the repo root)
import fs from 'node:fs';

const src = fs.readFileSync('dashboard.html', 'utf8');
const mod = src.match(/<script type="module">([\s\S]*?)<\/script>/)[1];

const START = 'function ccTitle(key, label, url)';
const END = 'async function hubSecPhotos()';
const i = mod.indexOf(START), j = mod.indexOf(END, i);
if (i < 0 || j < 0) { console.error('REGION NOT FOUND - markers moved'); process.exit(1); }
const region = mod.slice(i, j);
const recapRule = (mod.match(/const recapIsPublic = [^;]+;/) || [])[0];
if (!recapRule) { console.error('recapIsPublic not found'); process.exit(1); }

const escapeHtml = (v) => String(v == null ? '' : v)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const fmtDate = (d) => String(d).slice(0, 10);
const fmtDateTime = (d) => String(d).slice(0, 16).replace('T', ' ');
const mediaKindLabel = (u) => (/soundcloud/.test(u) ? 'SoundCloud audio' : 'YouTube video');
const SOCIAL_STAGES = ['idea', 'drafted', 'review', 'planned', 'scheduled', 'posted', 'archived'];
const SOCIAL_STAGE_LABEL = { idea: 'Idea', drafted: 'Drafted', review: 'In review', planned: 'Planned', scheduled: 'Scheduled', posted: 'Posted', archived: 'Archived' };
const toast = () => {};
const confirm = () => false;
const sb = { from: () => ({ update: () => ({ eq: async () => ({ error: null }) }), delete: () => ({ eq: async () => ({ error: null }) }) }) };

const hub = {
  eventId: 'e1',
  event: {
    id: 'e1', name: 'Henry Artist Showcase', is_featured: true,
    recap_videos: [
      { label: 'Full set', url: 'https://soundcloud.com/cw/full-set', is_public: true, artist_id: 'a1' },
      { label: 'Teaser', url: 'https://youtu.be/abcdefghijk', is_public: false },
    ],
  },
  _artists: [{ id: 'a1', display_name: 'KRNeY' }, { id: 'a2', display_name: 'Kloud9' }],
  _people: [],
  photos: [{ id: 'p1', is_public: true }, { id: 'p2', is_public: false }],
  assets: [
    { id: 'c1', kind: 'full', media: 'video', url: 'https://youtu.be/master1234', label: 'Master cut', duration_note: '58:00', artist_id: 'a1' },
    { id: 'c2', kind: 'clip', media: 'video', url: 'https://youtu.be/clip5678', label: 'IG cut', duration_note: '0:30' },
    { id: 'c3', kind: 'full', media: 'audio', url: 'https://soundcloud.com/cw/full-set', label: 'Already sent' },
    { url: 'https://soundcloud.com/cw/full-set', label: 'legacy', _legacy: true },
  ],
  eventPosts: [
    { id: 's1', title: 'Recap reel', stage: 'planned', scheduled_for: '2026-09-01T18:30:00Z', channels: ['instagram'], content_pillar: 'recap', link_url: 'https://instagram.com/p/x' },
    { id: 's2', title: 'Teaser', stage: 'idea', scheduled_for: null, channels: [] },
  ],
};

const fn = new Function(
  'escapeHtml,fmtDate,fmtDateTime,mediaKindLabel,SOCIAL_STAGES,SOCIAL_STAGE_LABEL,toast,confirm,sb,hub',
  recapRule + '\n' + region +
  '\n; return { hubRecapVideosHTML, hubAssetsHTML, hubPostsHTML, hubContentCardsHTML, hubUrlHost, hubRecapList };');
const api = fn(escapeHtml, fmtDate, fmtDateTime, mediaKindLabel, SOCIAL_STAGES, SOCIAL_STAGE_LABEL, toast, confirm, sb, hub);

let fails = 0;
const fail = (m) => { fails++; console.log('FAIL  ' + m); };
const pass = (m) => console.log('PASS  ' + m);
const render = (label, f) => {
  try {
    const out = f();
    const bad = out.match(/undefined|\[object Object\]|NaN/);
    if (bad) { fail(label + ' renders literal "' + bad[0] + '"'); return ''; }
    pass(label + ' renders (' + out.length + ' chars)');
    return out;
  } catch (e) { fail(label + ' threw: ' + e.message); return ''; }
};

const cards = render('summary cards', () => api.hubContentCardsHTML());
if (cards && !/On the site[\s\S]{0,90}>1</.test(cards)) fail('the "on the site" card should read 1');
else if (cards) pass('cards count live vs staged correctly');

// ---- the public list ---------------------------------------------------------
const rec = render('recap list', () => api.hubRecapVideosHTML());
for (const need of ['data-cc-rename="cv:0"', 'data-cv-artist="0"', 'data-cv-pub="0"', 'data-cv-del="0"']) {
  if (!rec.includes(need)) fail('the recap row is missing ' + need);
}
if (!fails) pass('recap rows: renameable name, artist tag, live/staged');
if (!/KRNeY[\s\S]{0,40}selected|selected[^>]*>KRNeY/.test(rec.replace(/\n/g, ''))) {
  if (!/value="a1" selected/.test(rec)) fail('the existing artist tag is not preselected');
}
if (!/value="a1" selected/.test(rec)) fail('artist a1 should be selected on the first row');
else pass('an existing artist tag comes back selected');
if (!/class="cc-todo"/.test(rec)) fail('the staged row is not banded as not-yet-live');
else pass('staged rows band amber, live rows green - the events-list vocabulary');

// ---- the library -------------------------------------------------------------
const lib = render('content library', () => api.hubAssetsHTML());
for (const need of ['data-cc-rename="ca:c1"', 'data-ca-field="kind"', 'data-ca-field="media"',
                    'data-ca-field="duration_note"', 'data-ca-field="artist_id"']) {
  if (!lib.includes(need)) fail('the library row cannot edit ' + need);
}
if (!fails) pass('library rows edit label, kind, media, duration and artist in place');
// The legacy pseudo-row is the recap list; it must not be duplicated here.
if (/legacy/.test(lib)) fail('the legacy recap entry is duplicated into the library');
else pass('the library excludes the recap list, so nothing is listed twice');
// An asset already on the public list offers no promote button.
if (!/data-ca-promote="c1"/.test(lib)) fail('a library asset not yet on the site has no "to the site" action');
if (/data-ca-promote="c3"/.test(lib)) fail('an asset already on the site list still offers to send it again');
else pass('"to the site" appears only where it would do something');

// ---- posts -------------------------------------------------------------------
const posts = render('social posts', () => api.hubPostsHTML());
for (const need of ['data-cc-rename="sp:s1"', 'data-sp-field="stage"', 'data-sp-field="scheduled_for"']) {
  if (!posts.includes(need)) fail('the post row cannot edit ' + need);
}
if (!posts.includes('data-hub-postedit="s1"')) fail('there is no way through to the full brief');
if (!/value="2026-09-01"/.test(posts)) fail('the scheduled date is not prefilled from the timestamp');
else pass('post rows edit title, stage and date, and keep a route to the full brief');
if (!/value="planned" selected/.test(posts)) fail('the current stage is not preselected');
else pass('the current stage comes back selected');

// ---- the write path must not flatten a timestamp -----------------------------
const patch = mod.slice(mod.indexOf('async function hubPostPatch'), mod.indexOf('async function hubPostPatch') + 800);
if (!/p\.scheduled_for\.slice\(11, 19\)/.test(patch)) {
  fail('editing the date drops the time - every post would move to midnight');
} else pass('editing the date keeps the time already on the post');


// ---- consistency with the events list ----------------------------------------
let cons = 0;
for (const [name, html] of [['recap', rec], ['library', lib], ['posts', posts]]) {
  if (!/<table class="data-table cc-table">/.test(html)) { fail(name + ' is not the events-list table'); cons++; }
  if (!/<thead><tr><th>/.test(html)) { fail(name + ' has no header row'); cons++; }
  if (!/class="cc-title-link" href="http/.test(html)) { fail(name + ' name is not a clickable link'); cons++; }
}
if (!cons) pass('all three use the events-list table, with header rows and clickable names');

if (/data-cv-label=|data-ca-field="label"|data-sp-field="title"/.test(rec + lib + posts)) {
  fail('a name is still a full-width input box');
} else pass('no name input boxes - renaming is behind the pencil');
if (!/data-cc-rename="cv:0"/.test(rec)) fail('there is no way to rename a recap row');
else pass('every row can still be renamed in place');

console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
