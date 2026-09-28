---
name: comewith-migration-number-collision
description: "Two migrations sharing a number silently overwrite each other's applied_migrations row — happened 2026-09-28, how it was repaired"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 21c71d6f-845a-4732-ab58-5d2a0b40974b
  modified: 2026-09-28T19:08:22.444Z
---

**In the Comewith repo, `git fetch` BEFORE choosing a migration number — always, even for a one-file change.** `MERGE_ROUTINE.md` step 0 says so and it is not ceremony.

**Why:** `public.applied_migrations` is keyed on **`version` alone**, and `db.py` upserts (`on conflict (version) do update set filename = excluded.filename`). So applying a second migration numbered `207` silently **overwrote the ledger row for `207_link_pages.sql`** — the record simply became a different file, with no warning. The objects from the clobbered migration stayed in prod; only the bookkeeping lied.

**What happened 2026-09-28:** local `master` was 22 commits stale (last local commit 2026-08-27). `ls supabase/migrations/ | tail` showed 206 as the highest, so `207_ingest_settlement.sql` looked free. `origin/master` already had 207–211. Checking the LOCAL directory is not checking; only `git fetch` then `git ls-tree -r --name-only origin/master supabase/migrations/` is.

**Repair that worked:** restore the clobbered row (`filename`, real `sha256` from `git show origin/master:<path>`, an inferred `applied_at` with the inference stated in `note`), renumber the new file to the next free number, insert its own row, and record the whole story in the migration's header comment. Verify the clobbered migration's objects still exist in prod first — if they do, only the record needs fixing, not the schema.

**Also: before `git checkout master` in that repo, check whether the incoming commits touch any locally-modified or deleted file.** Keith habitually carries uncommitted work there (receipts, render artefacts, thumbnails). Compare `git status --porcelain` against `git diff --name-only <old> origin/master` and only proceed on an empty intersection — never stash his work to get a rebase through.

**How to apply:** fetch first, pick the number from `origin/master`, and treat a stale local `master` as the default assumption rather than the exception.

Related: [[comewith-repo-and-push-auth]], [[cw-cash-reserve-inflows-bug]]
