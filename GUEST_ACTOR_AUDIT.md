# Guest / Actor Data-Hygiene Audit (read-only — FLAG only, nothing merged)

**Date:** 2026-06-16 · **Scope:** guests, actors, vendors (actor+vendor role), contractors, after the DI#2 ledger import.
**Rule:** this is a flag list. **No merge or standardization was executed** — auto-merging people on prod risks merging two different humans. Keith reviews; a follow-up acts on approved items.

## A. Same human as guest AND actor, but NOT linked (`guests.actor_id` null)
These guests share a name with an existing actor whose record has **no email**, so the email-based graduation couldn't connect them. Each is almost certainly one person in two lenses — suggest linking `guests.actor_id → actor.id` (and optionally copying the guest email onto the actor).
| Guest (email) | Actor | Suggested action |
|---|---|---|
| Adam Cohen (adam.ri.cohen@gmail.com) | Adam Cohen | link; set actor email |
| Ethan Pollak (epollak24@gmail.com) | Ethan Pollak | link (note: guest email ≠ the `pulse@pulsedevice.com` used to create the Pulse contact — confirm which is personal) |
| Francis Berkman (francisberkman@gmail.com) | Francis Berkman | link; set actor email |
| Liz McQuillan (emcquillan@gmail.com) | Liz McQuillan | link; set actor email |
| Patrick Savery (patrick.savery@gmail.com) | Patrick Savery | link; set actor email |
| Theresa Berkman (taberkman@gmail.com) | Theresa Berkman | link; set actor email |

## B. Possible duplicate actors / name variants (flagged during import — NOT created)
| Ledger name | Likely existing actor | Note |
|---|---|---|
| `Crossroads Cafe`, `Crossroads Cafe (Pastries)` | **Crossroads Café** | accent + suffix variant — same vendor; standardize to "Crossroads Café", attach the vendor lines |
| `Keith Berkman` | **Keith Berkman (Berky)** | same person; the import LINKED roles to Berky (no dup created) — confirm |
| `Teri Berkman` | **Theresa Berkman** | "Teri" likely nickname for Theresa — confirm same person |
| `Patrick Savery Sr.` | Patrick Savery | likely the **father** (different person) — confirm before any link |
| `Theresa McGuinness` | Theresa Berkman | matched on surname-ish only — **different person**, do not merge |
| `Lisa Meyer-Savery` | Patrick Savery | shares "Savery" — **different person**, do not merge |
| `Martin P. Salas` | Just Martin | matched on "Martin" only — **different person**, do not merge |
| `Amanda Brunidge` (ledger/door) | guest **Amanda Brundige** | spelling variant of the same attendee — standardize spelling |

## C. Low-quality / non-standard names
- **`Henry`** (actor, person, no surname, no email) — incomplete record; needs a real name.
- Single-token names that are **legitimate** (no action): `KRNeY`, `Kloud9`, `32LVS` (stage names); `BeWell`, `Freemind` (orgs).
- `Oheyitsamanda` (ledger donor) — an Instagram handle used as a name; flagged, not created.
- **Sara (Signal)** carries function `other` (from a prior sprint) — refine to a real function (booking/marketing) when known.

## D. Not duplicates (checked, no action)
- `Jennifer Alderman` vs `Jennifer Taveras` — shared first name only, **different people**. Fine.

## E. Artifacts that should never be people/actors (flagged, not created)
`Resident Advisor (RA) Payout`, `Zeffy Pending Payout`, `Postnet NY143 (Fliers)`, `Signal NYC (175 Morgan Ave LLC)` (the venue, not a vendor-person), `Anonymous Donor (Crossroads)` — payout/financial line items, not relationship-people.

## Summary
- Same-human unlinked: **6** (section A) — safe, high-confidence links awaiting approval.
- Possible-duplicate / variant: **8** (section B) — review before merge; 3 are likely *different* people, do not merge.
- Low-quality names: **2** real issues (Henry, Sara function) + handle-as-name.
- Artifacts excluded: **5**.

**Nothing in this report was executed.** Approve the ones you want and a follow-up will link/standardize them.
