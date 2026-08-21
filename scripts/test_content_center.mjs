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
if (!/class="cc-active"/.test(rec)) fail('the staged recap row is not banded');
else if (!/class="cc-done"/.test(rec)) fail('the live recap row is not banded');
else pass('recap: live green, staged blue');
// The words matter more than the colour - "how do I tell what is staged" was a
// fair question when the only clue was a dropdown value.
if (!/On the site<\/span>/.test(rec) || !/Staged<\/span>/.test(rec)) {
  fail('the recap row does not SAY whether it is staged, only shows a dropdown');
} else pass('every recap row says on-the-site or staged in words');

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


// ---- the three states a library asset can be in ------------------------------
// c1 is in the library only, c3 is on the public list and live.
if (!/not sent/.test(lib)) fail('a library-only asset does not say it has not been sent');
if (!/On the site/.test(lib)) fail('an asset that IS on the site does not say so');
if (!/<th>Where it is<\/th>/.test(lib)) fail('the library has no "where it is" column');
else pass('the library says where each asset stands, in words, in its own column');

// The header badge counts CONTENT, not photos. Frick Frack read 0 with a
// full-length mix in it because this counted hub.photos alone.
const tally = mod.slice(mod.indexOf('hub.counts.photos ='), mod.indexOf('hub.counts.photos =') + 320);
for (const part of ['hubRecapList()', 'hub.assets', 'hub.eventPosts', 'hub.photos.length']) {
  if (!tally.includes(part)) fail('the Content badge does not count ' + part);
}
if (!/full \\u00b7 |full · /.test(mod.slice(mod.indexOf("card('Library'"), mod.indexOf("card('Library'") + 260))) {
  fail('the Library card still describes everything as clips');
} else pass('the Library card counts full and clips separately');

// Promoting a short link has to resolve it first or the embed is dead on arrival.
const promo = mod.slice(mod.indexOf('async function hubAssetPromote'), mod.indexOf('async function hubAssetPromote') + 900);
if (!/resolveMediaUrls\(\[a\.url\]\)/.test(promo)) {
  fail('sending an asset to the site does not resolve the link - on.soundcloud.com short links will not embed');
} else pass('sending to the site resolves the link first');


// =============================================================================
// The content LIST - the Social Calendar's list view, rebuilt to the events-list
// spec (2026-08-21). It replaced the timeline. Same table, same banding, the
// four moving fields editable in place, chip filters instead of dropdowns.
// =============================================================================
const SOC_START = 'function socialFmtDate(d)';
const SOC_END = '// The filter strip lives in its own element';
const si = mod.indexOf(SOC_START), sj = mod.indexOf(SOC_END, si);
if (si < 0 || sj < 0) { console.error('SOCIAL REGION NOT FOUND - markers moved'); process.exit(1); }
const socRegion = mod.slice(si, sj);

const SOCIAL_CHANNELS = ['instagram', 'tiktok', 'facebook', 'x', 'youtube', 'email', 'blog', 'other'];
const SOCIAL_CHAN_LABEL = { instagram: 'Instagram', tiktok: 'TikTok', facebook: 'Facebook', x: 'X', youtube: 'YouTube', email: 'Email', blog: 'Blog', other: 'Other' };
const SOCIAL_STAGE_COLOR = { idea: '#8A7F72', posted: '#3DA35D' };
const social = {
  q: '', fStage: [], fSeries: [], fChan: [], view: 'list', selected: new Set(),
  noteCounts: { p1: 3 },
  posts: [
    { id: 'p1', title: 'Recap reel', stage: 'posted', posted_at: '2026-08-01T20:00:00Z',
      scheduled_for: '2026-07-31T18:30:00Z', channels: ['instagram', 'tiktok'], content_pillar: 'recap',
      series: 'Come With Parties', link_url: 'https://instagram.com/p/x',
      owner: { full_name: 'Janelle' }, event: { name: '7-11' } },
    { id: 'p2', title: 'Lineup drop', stage: 'scheduled', scheduled_for: '2026-09-04T17:00:00Z',
      channels: ['instagram'], content_pillar: 'lineup', series: 'Dance Infusion',
      owner: { email: 'liz@comewith.org' } },
    { id: 'p3', title: 'Studio session', stage: 'idea', scheduled_for: null, channels: [],
      content_pillar: 'takeover', series: null, caption: 'Berky in the booth' },
    { id: 'p4', title: 'Old announcement', stage: 'archived', posted_at: null,
      scheduled_for: '2026-01-02T12:00:00Z', channels: ['email'], series: 'Come With Parties' },
  ],
};
let filterHtml = '';
const $stub = (id) => (id === 'socialFilters'
  ? { set innerHTML(v) { filterHtml = v; }, get innerHTML() { return filterHtml; } }
  : null);

const socFn = new Function(
  'escapeHtml,fmtDate,fmtDateTime,mediaKindLabel,SOCIAL_STAGES,SOCIAL_STAGE_LABEL,SOCIAL_CHANNELS,SOCIAL_CHAN_LABEL,SOCIAL_STAGE_COLOR,toast,confirm,sb,hub,social,$',
  recapRule + '\n' + region + '\n' + socRegion +
  '\n; return { socialListHTML, socialFiltered, socialFilterDesc, renderSocialFilters, socialPillarList, socialChanCell };');
const soc = socFn(escapeHtml, fmtDate, fmtDateTime, mediaKindLabel, SOCIAL_STAGES, SOCIAL_STAGE_LABEL,
                  SOCIAL_CHANNELS, SOCIAL_CHAN_LABEL, SOCIAL_STAGE_COLOR, toast, confirm, sb, hub, social, $stub);

// ---- the table itself --------------------------------------------------------
const clist = render('content list', () => soc.socialListHTML(social.posts));
for (const need of ['data-sp-field="stage"', 'data-sp-field="scheduled_for"', 'data-sp-field="content_pillar"']) {
  if (!clist.includes(need)) fail('the list has no inline editor for ' + need);
}
if (!/data-sp-field="stage"[\s\S]{0,400}?<\/select>/.test(clist)) fail('the stage editor is not a select');
else pass('stage, date and pillar edit in place');
if (!/class="data-table cc-table"/.test(clist)) fail('the list is not the events-list table');
else pass('same data-table as the events list, with a real thead');
if (!/<thead><tr><th>Post<\/th>/.test(clist)) fail('the list has no real header row');

// The name is a LINK with the pencil beside it - never a full-width input box.
if (!/data-cc-rename="sl:p1"/.test(clist)) fail('there is no way to rename a post in place');
else pass('every row renames behind the pencil');
if (!/class="cc-title-link"/.test(clist)) fail('the post name is not a clickable link');
else pass('names are clickable links');
const stray = clist.match(/<input(?![^>]*type="date")[^>]*>/);
if (stray) fail('a name still renders as an input box: ' + stray[0].slice(0, 60));
else pass('no name input boxes on the list');

// ---- banding, on the events-list vocabulary ---------------------------------
const band = (id, cls) => {
  const row = clist.split('<tr class="').find(s => s.includes('sl:' + id));
  if (!row) { fail('row ' + id + ' is missing from the list'); return; }
  if (!row.startsWith(cls)) fail(id + ' should band ' + cls + ', got ' + row.slice(0, 9));
};
band('p1', 'cc-done');    // posted   -> green
band('p2', 'cc-active');  // scheduled-> blue
band('p3', 'cc-todo');    // idea     -> amber
band('p4', 'cc-off');     // archived -> muted
if (!fails) pass('three-colour banding: posted green, scheduled blue, idea amber, archived muted');

// ---- channels: an array, so a single select would delete data ---------------
const cell = soc.socialChanCell(social.posts[0]);
if (!/data-sp-chandel="p1" data-val="instagram"/.test(cell)) fail('a channel cannot be removed');
if (/<option value="instagram"/.test(cell)) fail('the add-a-channel list re-offers a channel the post already has');
if (!/<option value="youtube"/.test(cell)) fail('the add-a-channel list is missing an unused channel');
if (!/data-sp-chandel="p1" data-val="tiktok"/.test(cell)) fail('the second channel was dropped');
else pass('channels: both kept, one add select, no data loss');
if (!/value="2026-09-04"/.test(clist)) fail('the scheduled date does not reach the date input');
else pass('the scheduled day lands in the date box');

// content_pillar is FREE TEXT. A value nobody hardcoded must survive the select.
if (!/<option value="takeover" selected>/.test(clist)) {
  fail('a free-text pillar is not preselected - editing that row would erase it');
} else pass('an off-list pillar stays selected');
if (soc.socialPillarList().indexOf('takeover') < 0) fail('the pillar list is not derived from the data');

if (!/No posts match these filters\./.test(soc.socialListHTML([]))) {
  fail('an empty list says nothing');
} else pass('an empty list says why it is empty');

// ---- multi-select filtering --------------------------------------------------
const only = (patch) => {
  Object.assign(social, { q: '', fStage: [], fSeries: [], fChan: [] }, patch);
  const n = soc.socialFiltered().length;
  Object.assign(social, { q: '', fStage: [], fSeries: [], fChan: [] });
  return n;
};
const expect = (label, got, want) => (got === want ? pass(label) : fail(label + ': got ' + got + ', wanted ' + want));
expect('no filter means every post', only({}), 4);
expect('two stages at once', only({ fStage: ['idea', 'scheduled'] }), 2);
expect('series General means the ones tied to nothing', only({ fSeries: ['__none'] }), 1);
expect('two series at once', only({ fSeries: ['Come With Parties', 'Dance Infusion'] }), 3);
expect('a channel filter matches any channel on the post', only({ fChan: ['tiktok'] }), 1);
expect('two channels at once', only({ fChan: ['tiktok', 'email'] }), 2);
expect('groups combine (AND across, OR within)', only({ fStage: ['posted'], fChan: ['email'] }), 0);
expect('search reaches the caption', only({ q: 'booth' }), 1);
expect('search reaches the pillar', only({ q: 'takeover' }), 1);

social.fStage = ['idea', 'drafted'];
const desc = soc.socialFilterDesc();
if (!/All series/.test(desc) || !/Idea, Drafted/.test(desc)) fail('the shared filter description drops values: ' + desc);
else pass('the snapshot and email describe every selected value');
social.fStage = [];

// ---- the filter strip: chips, not dropdowns ---------------------------------
soc.renderSocialFilters();
if (/<select/.test(filterHtml)) fail('the filter strip still has a dropdown in it');
else pass('filters are chips, not single-value dropdowns');
for (const g of ['fStage', 'fSeries', 'fChan']) {
  if (!filterHtml.includes('data-scchip="' + g + '"')) fail('no chip group for ' + g);
}
if (!/data-val="__none"[^>]*>General/.test(filterHtml)) fail('there is no "General" chip for unfiled posts');
if (!/data-val="posted"[^>]*>Posted <span class="ev-chip-n">1<\/span>/.test(filterHtml)) {
  fail('chips do not carry their own count');
} else pass('every chip carries its count');
if (!/id="scCount">4 of 4</.test(filterHtml)) fail('the strip does not say how many of how many are shown');
else pass('the strip says 4 of 4');
if (/data-scclear/.test(filterHtml)) fail('"clear all" shows when nothing is filtered');
social.fChan = ['email'];
soc.renderSocialFilters();
if (!/data-scclear/.test(filterHtml)) fail('"clear all" is missing while a filter is on');
else pass('"clear all" appears only when there is something to clear');
if (!/class="ev-chip on" data-scchip="fChan" data-val="email"/.test(filterHtml)) {
  fail('the active chip is not marked on');
} else pass('the active chip reads as on');
social.fChan = [];

// The timeline is gone, and nothing may still reach for it.
for (const dead of ['timelineCardHtml', 'tl-card', "'timeline'"]) {
  if (mod.includes(dead)) fail('the timeline left ' + dead + ' behind');
}
if (!fails) pass('the timeline view is fully gone');


console.log(fails ? '\n' + fails + ' FAILURE(S)' : '\nAll checks passed.');
process.exit(fails ? 1 : 0);
