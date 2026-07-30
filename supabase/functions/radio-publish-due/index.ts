// radio-publish-due  (CRON-gated by ?secret=, no JWT)
//
// The scheduled release path. pg_cron used to call the SQL function
// public.radio_publish_due() directly, which is pure SQL and therefore CANNOT
// reach SoundCloud. That is why EP 1's mix embed was dead on arrival: the link
// was saved while the track was still private, the page flipped itself live on
// schedule, and nothing ever checked whether the embed actually worked.
//
// So before publishing anything, this VERIFIES the mix is embeddable and fixes it
// if it isn't:
//
//   1. oembed the stored mix_sc_track_url. If it answers, the link works as-is
//      and we touch nothing — that is the "or not do anything" case.
//   2. If oembed 404s, the track is private/scheduled. Resolve it through the
//      OAuth'd account and flip sharing=public, then re-check.
//   3. If there is no URL at all, find the upload on the account by runtime/title
//      (same logic as sc-connect's find_mix) and store it.
//   4. THEN publish, by calling the SQL function that does the DB side.
//
// Publishing is never blocked by a SoundCloud problem — the drop goes out and the
// failure is recorded on the station instead of vanishing. Silent failure is the
// thing this function exists to stop.
//
// GET/POST ?secret=<RADIO_PUBLISH_SECRET>[&dry=1][&station=N]
//   dry=1      report what it WOULD do, publish nothing
//   station=N  only consider that station (ignores the due filter, for testing)
import { createClient } from "npm:@supabase/supabase-js@2";

const JH = { "Content-Type": "application/json" };
const SC_ACCEPT = "application/json; charset=utf-8";
const UA = "Mozilla/5.0 (compatible; ComeWithRadio/1.0)";

const ok = (b: unknown) => new Response(JSON.stringify(b), { headers: JH });
const err = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

// Does the public site's player have any chance of rendering this? oembed is the
// exact call the embed makes, so its answer is the ground truth — a 200 page does
// NOT imply an embeddable track.
async function embeddable(url: string): Promise<boolean> {
  try {
    const r = await fetch("https://soundcloud.com/oembed?format=json&url=" + encodeURIComponent(url),
      { headers: { "User-Agent": UA } });
    return r.ok;
  } catch { return false; }
}

Deno.serve(async (req) => {
  const u = new URL(req.url);
  const secret = Deno.env.get("RADIO_PUBLISH_SECRET");
  if (!secret || u.searchParams.get("secret") !== secret) return err(401, "bad secret");

  const dry = u.searchParams.get("dry") === "1";
  const onlyStation = u.searchParams.get("station");
  const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);

  let sel = admin.from("sc_playlists")
    .select("id, station_no, name, slug, status, published, scheduled_go_live, mix_sc_track_url, mix_sc_track_id");
  if (onlyStation) sel = sel.eq("station_no", Number(onlyStation));
  else sel = sel.not("scheduled_go_live", "is", null).lte("scheduled_go_live", new Date().toISOString())
                .neq("status", "live").is("published", false);
  const { data: due, error: selErr } = await sel;
  if (selErr) return err(500, selErr.message);
  if (!due?.length) return ok({ success: true, due: 0, note: "nothing scheduled is due" });

  // SoundCloud OAuth, refreshed when stale (refresh tokens are single-use).
  const clientId = Deno.env.get("SC_CLIENT_ID");
  const clientSecret = Deno.env.get("SC_CLIENT_SECRET");
  const { data: oauth } = await admin.from("sc_oauth").select("*").eq("id", "singleton").maybeSingle();
  async function scToken(): Promise<string | null> {
    if (!oauth?.access_token) return null;
    const stale = oauth.expires_at && new Date(oauth.expires_at).getTime() < Date.now() + 60_000;
    if (!stale || !oauth.refresh_token || !clientId || !clientSecret) return oauth.access_token;
    const rr = await fetch("https://secure.soundcloud.com/oauth/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", "accept": SC_ACCEPT },
      body: new URLSearchParams({ grant_type: "refresh_token", client_id: clientId,
        client_secret: clientSecret, refresh_token: oauth.refresh_token }),
    });
    const rt = await rr.json().catch(() => ({}));
    if (rr.ok && rt.access_token) {
      await admin.from("sc_oauth").update({
        access_token: rt.access_token, refresh_token: rt.refresh_token || oauth.refresh_token,
        expires_at: new Date(Date.now() + (rt.expires_in || 3600) * 1000).toISOString(),
        updated_at: new Date().toISOString(),
      }).eq("id", "singleton");
      return rt.access_token;
    }
    return null;
  }

  const results: any[] = [];
  for (const pl of due) {
    const step: any = { station_no: pl.station_no, sc: "untouched", published: false, warning: null };
    let url: string | null = pl.mix_sc_track_url || null;
    let trackId: string | null = pl.mix_sc_track_id || null;

    // 1 · already fine?
    if (url && await embeddable(url)) {
      step.sc = "already embeddable";
    } else {
      const token = await scToken();
      if (!token) {
        step.warning = "SoundCloud not connected — mix left as-is.";
      } else {
        // 2 · we have a URL but it won't embed: it is private. Flip it public.
        if (url) {
          // /resolve does NOT return a PRIVATE track from its plain permalink, so
          // resolve first and fall back to scanning the account's own uploads for a
          // matching permalink — /me/tracks is the only view that includes privates.
          const rr = await fetch("https://api.soundcloud.com/resolve?url=" + encodeURIComponent(url),
            { headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT } });
          let rj = await rr.json().catch(() => ({}));
          let found = rr.ok && rj?.kind === "track";
          if (!found) {
            const bare = (x: string) => (x || "").split("?")[0].replace(/\/+$/, "").toLowerCase();
            const lr = await fetch("https://api.soundcloud.com/me/tracks?limit=50&linked_partitioning=true",
              { headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT } });
            const lj = await lr.json().catch(() => ({}));
            const mine: any[] = Array.isArray(lj) ? lj : (lj.collection || []);
            const m = mine.find((t) => bare(t.permalink_url) === bare(url!));
            if (m) { rj = m; found = true; step.sc = "found via /me/tracks (private)"; }
            else if (!lr.ok) step.warning = `SoundCloud wouldn't list your uploads (${lr.status}).`;
          }
          if (found) {
            trackId = String(rj.id);
            if (rj.sharing === "private" && !dry) {
              const fd = new FormData();
              fd.set("track[sharing]", "public");
              const up = await fetch(`https://api.soundcloud.com/tracks/${rj.id}`, {
                method: "PUT", headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT }, body: fd });
              const uj = await up.json().catch(() => ({}));
              if (up.ok) {
                url = typeof uj.permalink_url === "string" ? uj.permalink_url : url;
                step.sc = (await embeddable(url!)) ? "flipped public — embeddable" : "flipped public — oembed still cold";
              } else {
                step.warning = `Could not make the track public (${up.status}).`;
                step.sc = "still private";
              }
            } else {
              step.sc = dry ? `would flip public (sharing=${rj.sharing})` : `sharing=${rj.sharing}`;
            }
          } else {
            step.warning = "Stored mix link didn't resolve on your account.";
          }
        } else {
          // 3 · no URL at all: find the upload on the account.
          const lr = await fetch("https://api.soundcloud.com/me/tracks?limit=50&linked_partitioning=true",
            { headers: { "Authorization": "OAuth " + token, "accept": SC_ACCEPT } });
          const lj = await lr.json().catch(() => ({}));
          const mine: any[] = Array.isArray(lj) ? lj : (lj.collection || []);
          const norm = (x: string) => (x || "").toLowerCase().replace(/[^a-z0-9]/g, "");
          const keys = [norm(pl.name), norm(pl.slug || ""), "ep" + (pl.station_no ?? "")].filter((k) => k.length > 2);
          const hit = mine.map((t) => {
            const nt = norm(t.title);
            let sc = 0;
            for (const k of keys) if (k && (nt.includes(k) || k.includes(nt))) sc += 0.5;
            return { t, sc };
          }).sort((a, b) => b.sc - a.sc)[0];
          if (hit && hit.sc >= 0.5) {
            trackId = String(hit.t.id); url = hit.t.permalink_url;
            step.sc = dry ? `would link "${hit.t.title}"` : `linked "${hit.t.title}"`;
          } else {
            step.warning = "No mix link set and no upload on the account matched this episode.";
          }
        }
      }
    }

    if (!dry) {
      if (url && (url !== pl.mix_sc_track_url || trackId !== pl.mix_sc_track_id)) {
        await admin.from("sc_playlists").update({
          mix_sc_track_url: url, mix_sc_track_id: trackId, updated_at: new Date().toISOString(),
        }).eq("id", pl.id);
      }
      // 4 · publish regardless — a SoundCloud problem must never hold the drop.
      const { data: slug, error: pubErr } = await admin.rpc("radio_publish_station", { p_id: pl.id });
      step.published = !pubErr && !!slug;
      step.slug = slug || null;
      if (pubErr) step.warning = (step.warning ? step.warning + " | " : "") + "publish failed: " + pubErr.message;
      // Record the warning where a human will see it, rather than losing it to logs.
      if (step.warning) {
        await admin.from("station_notes").insert({
          station_id: pl.id, body: "⚠️ Scheduled publish: " + step.warning,
        }).then(() => {}, () => {});
      }
    }
    results.push(step);
  }
  return ok({ success: true, dry, due: due.length, results });
});
