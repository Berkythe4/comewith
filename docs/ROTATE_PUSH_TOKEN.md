# Rotate the push token

The token Jennifer uses to POST finance rows to `ingest-finance`. Rotation is the
real control here — the token is long-lived and static, so being able to replace
it quickly and without drama is what keeps that acceptable.

Takes about two minutes. Nothing below needs a deploy.

## Rotate

**1. Generate a new value** (on your machine, in a normal terminal — not through
an assistant, and not in a session transcript):

```
echo "push_$(openssl rand -hex 32)"
```

**2. Set it on the SERVER first.**

```
supabase secrets set PUSH_TOKEN=push_<new value>
```

Server first, always. If Jennifer changes first, every push 401s until the server
catches up. This way the only casualty is pushes still using the old value, and
step 4 fixes that immediately.

> **Zero-downtime note.** The endpoint accepts exactly one token, so there is a
> brief window between step 2 and step 4 where Jennifer's pushes fail with 401.
> That is fine — the import is manual and re-runnable, and a failed push loses
> nothing. If that ever stops being true, add a second accepted secret
> (`PUSH_TOKEN_PREVIOUS`) and check both, then clear it after the window.

**3. Update Jennifer.** Edit `PUSH_TOKEN` in the planner's own `.env` on Keith's
machine. Not this repo — this repo never holds the value.

**4. Verify one push.** Run an import in Jennifer and confirm the summary panel
reports the rows as sent. A `401` means step 2 and step 3 disagree.

**5. Nothing to remove.** Setting the secret in step 2 replaced the old value.
If you added `PUSH_TOKEN_PREVIOUS` for a window, unset it now:

```
supabase secrets unset PUSH_TOKEN_PREVIOUS
```

## When to rotate

- **Immediately** on any suspected exposure — pasted into a chat, committed and
  pushed, sent over email or Slack, read aloud on a call, or captured in a
  screen recording.
- **Immediately** if `git log --all --full-history -- .env` ever returns a commit.
  That means the value reached history; rotating is step one and rewriting
  history is a separate, bigger decision — raise it with Keith.
- On offboarding, whenever someone with repo or machine access moves on.
- **At least annually**, even when nothing looks wrong. A token nobody has ever
  rotated is a token nobody knows how to rotate.

## If you think it leaked

1. Rotate first (above). Do not investigate first — rotation is cheap and
   instant, investigation is neither.
2. Then work out the blast radius. The token is **write-only to one endpoint**:
   it can push fee and vendor rows, and nothing else. It grants no read access,
   no database access, and no ability to reach any other function. It is not a
   Supabase key.
3. Check the function's logs for pushes you cannot account for.

## What this token is not

It is not the Supabase service-role key, not `SBP_PAT`, and not a Resend key.
Those are far more powerful and rotate differently. If one of *those* leaked,
this runbook is the wrong document.
