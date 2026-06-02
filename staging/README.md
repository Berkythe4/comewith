# /staging — admin-gated review area

Password-free, **reusable** gate for review-before-publish pages. Reuses the **same
Supabase sign-in as `dashboard.html`** (no second password system). One front door:
sign in once at `/dashboard.html`, and every `/staging/` page sees the session.

## How the gate works

`guard.js` is the only place auth is wired. It:
1. Creates the Supabase client (same project URL + publishable key as the dashboard).
2. `getSession()` — no session → redirects to `/dashboard.html` (the existing sign-in).
3. Reads `profiles.role` — `master_admin` / `sub_admin` (the `is_admin()` set) → reveals
   the page; anyone else → "admins only" notice.

It's **client-side gating on a static host** — enough to keep pages out of public/casual
view, **not real security**. Keep genuinely sensitive data (financials, rosters, venues)
in Supabase behind RLS, **not** as static files here. Staging is for review pages only.

## Add a new report (2 steps)

**1. Gate the page.** Put these two lines in the page's `<head>`, as early as possible
(first line prevents a flash; second runs the shared guard):

```html
<script>document.documentElement.style.visibility='hidden'</script>
<script type="module" src="/staging/guard.js"></script>
```

That's the entire auth wiring — no per-page config. To publish a page publicly later,
just delete those two lines.

**2. List it in the hub.** Add one entry to the `REPORTS` array in `staging/index.html`:

```js
{ title: "My Report", desc: "One-line description.",
  href: "/path/to/report.html", status: "review" }
```

A page can live **anywhere** under the site root (it doesn't have to be inside
`/staging/`) — the gate works by absolute path `/staging/guard.js`. Self-contained pages
can also just be dropped directly into `/staging/`.

## Current contents
- `guard.js` — shared admin gate (the only auth wiring).
- `index.html` — gated hub listing available reports (edit the `REPORTS` array to add).
- First items (gated in place, paths/deps intact):
  - `/events/dance-infusion/di-02-2026-05/reports/impact-report.html`
  - `/events/dance-infusion/di-02-2026-05/reports/public-audit.html`

## URLs (after deploy)
- Hub: `https://comewith.org/staging/`
- Impact report: `https://comewith.org/events/dance-infusion/di-02-2026-05/reports/impact-report.html`
- Public audit: `https://comewith.org/events/dance-infusion/di-02-2026-05/reports/public-audit.html`

Not signed in → any of these bounce to `/dashboard.html` to sign in, then are reachable.
