// node --test supabase/functions/sc-connect/radio_post.test.ts
//
// Go live's release card must be born classified (phase radio, account
// come_with, format reel) and must keep everything it carried before 220.
// The SQL twin is tested in supabase/tests/220_radio_post_classified.rollback_test.sql.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { radioReleasePost } from "./radio_post.ts";

const NOW = "2026-10-08T23:00:00.000Z";

test("a Go live release card is phase radio, Come With, reel", () => {
  const p = radioReleasePost({ station_no: 11, name: "Come With NYC Radio Ep6" }, "Miss Vee takes over", NOW, "https://comewith.org/radio.html?s=x");
  assert.equal(p.phase, "radio");
  assert.equal(p.account, "come_with");
  assert.equal(p.format, "reel");
});

test("everything it carried before is unchanged", () => {
  const p = radioReleasePost({ station_no: 11, name: "Come With NYC Radio Ep6" }, "desc", NOW, "https://comewith.org/radio.html?s=x");
  assert.equal(p.title, "📻 Come With Radio SHOW 11 — Come With NYC Radio Ep6");
  assert.deepEqual(p.channels, ["other"]);
  assert.equal(p.series, "Come With Radio");
  assert.equal(p.content_pillar, "radio episode");
  assert.equal(p.stage, "posted");
  assert.equal(p.scheduled_for, NOW);
  assert.equal(p.posted_at, NOW);
  assert.equal(p.link_url, "https://comewith.org/radio.html?s=x");
  assert.equal(p.caption, "desc");
});

test("empty description -> null caption, long one capped at 1000", () => {
  assert.equal(radioReleasePost({ station_no: 1, name: "x" }, "", NOW, "u").caption, null);
  assert.equal(radioReleasePost({ station_no: 1, name: "x" }, "a".repeat(1500), NOW, "u").caption!.length, 1000);
});

test("sc-connect's finalize inserts exactly this payload", () => {
  const src = readFileSync(new URL("./index.ts", import.meta.url), "utf8");
  assert.match(src, /from\("social_posts"\)\.insert\(radioReleasePost\(/);
  assert.equal((src.match(/from\("social_posts"\)\.insert\(/g) || []).length, 1, "no second, unclassified insert");
});

test("the SQL twin (migration 220) sets the same three values", () => {
  const sql = readFileSync(new URL("../../migrations/220_radio_post_classified.sql", import.meta.url), "utf8");
  assert.match(sql, /link_url, phase, account, format\)/);
  assert.match(sql, /'radio', 'come_with', 'reel'\);/);
});
