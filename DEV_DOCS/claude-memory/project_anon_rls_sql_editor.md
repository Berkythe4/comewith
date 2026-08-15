---
name: project-anon-rls-sql-editor
description: "SUPERSEDED 2026-05-28 in Phase 6. The \"anon RLS bug\" was actually a PostgREST RETURNING quirk — INSERT-with-return-representation triggers a SELECT-after-INSERT that fails because anon has no SELECT policy on inquiries. Use `return=minimal` (or just don't chain .select())."
metadata: 
  node_type: memory
  type: project
  originSessionId: d1806c45-1082-43f8-96a4-06579768d931
---

# RESOLVED — keep for the lesson, not the workaround

## The real cause (discovered Phase 6 sprint 1)

`INSERT ... RETURNING ...` and PostgREST's default `Prefer: return=representation`
both require Postgres to ALSO evaluate SELECT RLS on the inserted row. If the
anon role has no SELECT policy (which is the case for `public.inquiries` — only
admins can read), the post-insert SELECT fails. Postgres reports this back as
`42501: new row violates row-level security policy for table "..."` — a
misleading error because the INSERT WITH CHECK itself succeeded.

This was repeatable across:
- SQL Editor: `set local role anon; insert ... returning ...;`
- PostgREST: POST with default `Prefer: return=representation`
- supabase-js: `.from(t).insert(...).select(...).single()` (the `.select()`
  chain sends `Prefer: return=representation`)

And it disappeared when:
- `Prefer: return=minimal` header was sent → HTTP 201
- A single `FOR ALL` PERMISSIVE policy replaced the per-cmd policies, because
  the FOR ALL USING clause incidentally granted anon SELECT

## How to apply

For any anon-write to a table where anon has no SELECT policy:
- Use `Prefer: return=minimal` for raw PostgREST calls
- For supabase-js, drop any `.select()` chain from the insert
- Don't use `INSERT ... RETURNING` in raw SQL run as anon
- OR add a narrow SELECT policy if the client actually needs the row back

The simplest path is just to NOT ask for the row back when anon writes. The
client doesn't need it — confirmation that the insert didn't error is enough.

## Why we believed the wrong thing for a while

Phase 2 debugging never tried `Prefer: return=minimal` and never separated
"INSERT didn't happen" from "INSERT happened but SELECT-after failed." The
error message is identical in both cases. Lesson: when you see 42501 on a
write that has a `RETURNING` or `Prefer: return=representation`, suspect a
missing SELECT policy before suspecting a missing WITH CHECK.
