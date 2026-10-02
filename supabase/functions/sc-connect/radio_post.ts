// The "posted" release card that Go live drops on the social calendar.
//
//   node --test supabase/functions/sc-connect/radio_post.test.ts
//
// Kept apart from index.ts so it can be tested, and so the classification is
// written down once on this side. The scheduled go-live builds the same card
// in SQL (radio_publish_station, migration 220) - change one, change the other.
// A release is the episode video, so it is a reel; recaps are carousels and are
// planned by hand or by the Claude connector, never auto-created here.

export type ReleaseStation = { station_no: number | null; name: string | null };

export function radioReleasePost(pl: ReleaseStation, descSc: string | null, nowIso: string, pageUrl: string) {
  return {
    title: `📻 Come With Radio SHOW ${pl.station_no ?? ""} — ${pl.name || ""}`.trim(),
    caption: (descSc || "").slice(0, 1000) || null,
    channels: ["other"], series: "Come With Radio", content_pillar: "radio episode",
    stage: "posted", scheduled_for: nowIso, posted_at: nowIso,
    link_url: pageUrl,
    phase: "radio", account: "come_with", format: "reel",
  };
}
