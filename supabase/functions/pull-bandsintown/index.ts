// pull-bandsintown
//
// The first ARTIST-FIRST source. RA, DICE and Ticketmaster are all ticketing
// marketplaces: they answer "what is on sale in this city", so an artist selling
// direct is invisible to every one of them and a fourth marketplace would not
// change that. Bandsintown answers a different question — "where is this artist
// playing" — from dates the artist posts themselves, so it sees a show whoever
// sells the ticket. That is what closes the Lane 8 gap.
//
// ── HOW "ELECTRONIC ONLY" IS GUARANTEED ─────────────────────────────────────
// NOT by filtering what comes back. Bandsintown's event payload carries no genre
// at all, so any genre rule here would be invented. The guarantee comes from the
// INPUT: this function only ever asks about artists WE already track — the
// watchlist, partners, and the ra_artists pool that RA/DICE/TM built. Charli XCX
// cannot arrive through this door because we never ask Bandsintown about her.
// Widening the input is therefore the only way to let a non-electronic act in,
// which makes the decision explicit and reviewable instead of a threshold.
//
// The one thing that does come back un-vetted is the rest of the BILL: an
// artist's event lists its whole lineup, and a support act may be anything. That
// is accurate data about a real show and is stored as such — but a name only
// enters the ARTIST POOL if we already knew it. See the ra_artists note below.
//
// ── SECRET ──────────────────────────────────────────────────────────────────
// BANDSINTOWN_APP_ID. Missing = a documented NO-OP with a reason, never a silent
// error and never a fabricated empty result.
//
// Body: { scope?: "watchlist" | "partners" | "pool"   (default "watchlist")
//         names?: string[]        // ad-hoc, overrides scope
//         days?: number           // default 60
//         offset?: number         // rotation start, see the wall-clock note
//         limit?: number }        // max artists this run (default 40)
//
// Admin JWT OR service-role bearer, the pattern pull-ra-market already uses.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JH = { ...CORS, "Content-Type": "application/json" };
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

// The edge runtime stops at 150s. Leave room to write what we have: a run that
// dies mid-flight would otherwise leave no trace of the artists it did read.
const WALL_MS = 105_000;

// NYC by COORDINATES, not by city name. Ticketmaster taught this one: matching
// "New York" literally returned Manhattan only and Brooklyn Steel / Avant
// Gardner did not exist as far as that pull was concerned. Bandsintown returns
// lat/lng on every venue, and a box does not care what the promoter typed.
const NYC = { latMin: 40.45, latMax: 40.95, lngMin: -74.30, lngMax: -73.65 };
const inNyc = (v: Record<string, unknown>) => {
  const la = Number(v?.latitude), ln = Number(v?.longitude);
  if (Number.isFinite(la) && Number.isFinite(ln)) {
    return la >= NYC.latMin && la <= NYC.latMax && ln >= NYC.lngMin && ln <= NYC.lngMax;
  }
  // No coordinates: fall back to the city name rather than dropping the show.
  const city = String(v?.city ?? "").trim().toLowerCase();
  const region = String(v?.region ?? "").trim().toLowerCase();
  return (region === "ny" || region === "new york") &&
    ["new york", "brooklyn", "queens", "bronx", "the bronx", "staten island",
     "long island city", "ridgewood", "astoria"].includes(city);
};

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return err(405, "POST only");

  const url = Deno.env.get("SUPABASE_URL")!;
  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
  const auth = req.headers.get("Authorization") ?? "";
  const token = auth.replace(/^Bearer\s+/i, "");
  if (!token) return err(401, "Missing bearer token");

  const admin = createClient(url, serviceKey, { auth: { persistSession: false } });

  // Service-role bearer OR an admin JWT.
  if (token !== serviceKey) {
    const { data: u } = await admin.auth.getUser(token);
    if (!u?.user) return err(401, "Not signed in");
    const { data: prof } = await admin.from("profiles").select("role, deleted_at")
      .eq("id", u.user.id).maybeSingle();
    if (!prof || prof.deleted_at || !["master_admin", "sub_admin"].includes(String(prof.role))) {
      return err(403, "Admins only");
    }
  }

  const appId = Deno.env.get("BANDSINTOWN_APP_ID");
  if (!appId) {
    // A documented no-op. It must not read as "Bandsintown had nothing".
    return new Response(JSON.stringify({
      success: false, skipped: true, source: "bandsintown",
      reason: "BANDSINTOWN_APP_ID is not set — nothing was requested. Set the secret, then run this again.",
    }), { headers: JH });
  }

  try {
    const b = await req.json().catch(() => ({})) as Record<string, unknown>;
    const days = Number.isFinite(Number(b.days)) ? Math.max(1, Math.min(365, Number(b.days))) : 60;
    const limit = Number.isFinite(Number(b.limit)) ? Math.max(1, Math.min(200, Number(b.limit))) : 40;
    const offset = Number.isFinite(Number(b.offset)) ? Math.max(0, Number(b.offset)) : 0;
    const scope = ["watchlist", "partners", "pool"].includes(String(b.scope)) ? String(b.scope) : "watchlist";

    const today = new Date();
    const from = today.toISOString().slice(0, 10);
    const to = new Date(today.getTime() + days * 86400000).toISOString().slice(0, 10);

    // ── Who to ask about. This list IS the genre policy. ──────────────────
    let names: string[] = [];
    if (Array.isArray(b.names) && b.names.length) {
      names = b.names.map((n) => String(n).trim()).filter(Boolean);
    } else if (scope === "watchlist") {
      const { data } = await admin.from("watchlist").select("label")
        .eq("kind", "artist").or("archived.is.null,archived.eq.false");
      names = (data ?? []).map((r) => String(r.label ?? "").trim()).filter(Boolean);
    } else if (scope === "partners") {
      const { data } = await admin.from("ra_artists").select("name").eq("is_partner", true);
      names = (data ?? []).map((r) => String(r.name ?? "").trim()).filter(Boolean);
    } else {
      const { data } = await admin.from("ra_artists").select("name")
        .gte("next_event_date", from).order("name");
      names = (data ?? []).map((r) => String(r.name ?? "").trim()).filter(Boolean);
    }
    // De-dupe case-insensitively, keep a stable order so rotation is meaningful.
    const seen = new Set<string>();
    names = names.filter((n) => { const k = n.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; })
      .sort((a, c) => a.localeCompare(c));

    const total = names.length;
    // ROTATE. Starting at #1 every run means the tail is never read however often
    // the button is pressed — a permanent blind spot that reports itself clean.
    const start = total ? offset % total : 0;
    const ordered = total ? names.slice(start).concat(names.slice(0, start)) : [];
    const batch = ordered.slice(0, limit);

    const began = Date.now();
    const rows: Record<string, unknown>[] = [];
    const artistRows: Record<string, unknown>[] = [];
    const perArtist: Record<string, number> = {};
    const notFound: string[] = [];
    const failed: string[] = [];
    let skippedForTime: string[] = [];

    // Names we already track — a name on someone else's bill only enters the
    // ARTIST POOL if we already knew it. Storing the bill is accurate; promoting
    // an unknown support act into the pool would import whatever genre it is.
    const { data: knownRows } = await admin.from("ra_artists").select("name");
    const known = new Set((knownRows ?? []).map((r) => String(r.name ?? "").trim().toLowerCase()));

    for (let i = 0; i < batch.length; i++) {
      if (Date.now() - began > WALL_MS) { skippedForTime = batch.slice(i); break; }
      const name = batch[i];
      let json: unknown;
      try {
        const r = await fetch(
          `https://rest.bandsintown.com/artists/${encodeURIComponent(name)}/events` +
          `?app_id=${encodeURIComponent(appId)}&date=upcoming`,
          { headers: { Accept: "application/json" }, signal: AbortSignal.timeout(12_000) },
        );
        // Validate the PAYLOAD, not the status. Bandsintown answers 200 with an
        // error object for an unknown artist and for a rejected app_id alike —
        // reading r.ok alone would report every artist as "no shows".
        json = await r.json().catch(() => null);
        if (!Array.isArray(json)) {
          const msg = JSON.stringify(json ?? "").slice(0, 200);
          if (/app_id|unauthor|forbid|invalid/i.test(msg)) {
            return new Response(JSON.stringify({
              success: false, source: "bandsintown",
              reason: "Bandsintown rejected the app_id — check the BANDSINTOWN_APP_ID secret.",
              detail: msg, asked: i,
            }), { headers: JH });
          }
          notFound.push(name);
          continue;
        }
      } catch (_e) {
        failed.push(name);
        continue;
      }

      let kept = 0;
      for (const ev of json as Record<string, unknown>[]) {
        const venue = (ev?.venue ?? {}) as Record<string, unknown>;
        const dt = String(ev?.datetime ?? "").slice(0, 10);
        if (!dt || dt < from || dt > to) continue;
        if (!inNyc(venue)) continue;

        const lineup = Array.isArray(ev?.lineup) ? (ev.lineup as unknown[]).map(String) : [name];
        const offers = Array.isArray(ev?.offers) ? ev.offers as Record<string, unknown>[] : [];
        const ticket = offers.find((o) => String(o?.type ?? "").toLowerCase() === "tickets");
        rows.push({
          ra_id: "bit:" + String(ev?.id ?? `${name}-${dt}`).slice(0, 60),
          title: String(ev?.title ?? "").trim() || lineup.join(", "),
          event_date: dt,
          venue_name: String(venue?.name ?? "").trim() || null,
          lineup: lineup.map((n) => ({ name: n })),
          source: "bandsintown",
          content_url: String(ticket?.url ?? ev?.url ?? "") || null,
          is_ticketed: !!ticket,
          // attending stays NULL: Bandsintown has no RSVP count, and a 0 would be
          // read by the buzz score as "nobody is going".
          fetched_at: new Date().toISOString(),
        });
        kept++;

        // The artist we asked about enters the pool if they were not in it — that
        // is how Lane 8 becomes visible. Support acts do not (see `known`).
        const key = name.toLowerCase();
        if (!known.has(key)) {
          known.add(key);
          artistRows.push({
            ra_id: "bit_" + key.replace(/[^a-z0-9]+/g, "-").slice(0, 50),
            name, source: "bandsintown", is_partner: false,
            next_event_date: dt,
            next_venue: String(venue?.name ?? "").trim() || null,
            next_event_title: String(ev?.title ?? "").trim() || null,
            next_event_url: String(ticket?.url ?? ev?.url ?? "") || null,
            fetched_at: new Date().toISOString(),
          });
        }
      }
      if (kept) perArtist[name] = kept;
    }

    // UPSERT ONLY — deliberately no delete keyed to the window. This run reads a
    // ROTATING SUBSET of the artists, so deleting `source='bandsintown'` across
    // the date range would throw away shows for every artist this run did not ask
    // about. Past rows are pruned instead, which is always safe.
    if (rows.length) {
      const { error } = await admin.from("ra_events").upsert(rows, { onConflict: "ra_id" });
      if (error) return err(500, "Could not save Bandsintown events: " + error.message);
    }
    if (artistRows.length) {
      const { error: ae } = await admin.from("ra_artists").upsert(artistRows, { onConflict: "ra_id" });
      if (ae) console.error("bandsintown ra_artists:", ae.message);
    }
    await admin.from("ra_events").delete().eq("source", "bandsintown").lt("event_date", from);

    const asked = batch.length - skippedForTime.length;
    return new Response(JSON.stringify({
      success: true,
      source: "bandsintown",
      scope,
      // PARTIAL is its own state, and it names what it missed rather than a count:
      // "we read 40 of 300" with no names is a clean-looking report of a gap.
      status: skippedForTime.length ? "PARTIAL" : "OK",
      artists_total: total,
      asked,
      next_offset: total ? (start + asked) % total : 0,
      saved: rows.length,
      new_artists: artistRows.length,
      with_shows: perArtist,
      no_shows_or_unknown: notFound,
      request_failed: failed,
      skipped_for_time: skippedForTime,
      from, to,
      note: "Electronic-only is enforced by WHO we ask about (watchlist / partners / the existing pool), not by filtering the response — Bandsintown returns no genre.",
    }), { headers: JH });
  } catch (e) {
    return err(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
