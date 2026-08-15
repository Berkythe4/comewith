---
name: project-beatport-cart-api
description: "Verified Beatport v4 cart API shape (add-to-cart works) — discovered 2026-07-29 against Keith's live account"
metadata: 
  node_type: memory
  type: project
  originSessionId: 279f1814-6e4e-4f39-a354-6447da747ad9
  modified: 2026-07-29T21:14:10.188Z
---

Beatport's cart API **does** work with a pasted user token (`scope: app:prostore user:dj`).
Verified end-to-end 2026-07-29 by filling station 2's cart (13 tracks added, confirmed
by re-reading the cart). Supersedes the "NO purchase API" note on the APIs map.

- Discover endpoints from `GET /v4/` → `{catalog, curation, my}` and `GET /v4/my/` →
  lists `carts`, `default-cart`, `default-cart/items`, `downloads`, `account`, …
  **`/v4/my/cart/` (singular) 404s** — the earlier dead end.
- `GET /v4/my/carts/` → `[{default:true, id:<cartId>, name:"cart"}, {name:"hold-bin"}]`.
  Keith's default cart id = **284164252**.
- `GET /v4/my/carts/<cartId>/items/` → `{"tracks":[…]}`; `GET /v4/my/carts/<cartId>/`
  → `{"releases":[…]}`. **Tracks and releases are separate lists** — a release added
  from a release page never shows up in `tracks`, so count both before claiming a total.
- Add: `POST /v4/my/carts/<cartId>/items/` with
  `{item_id:<trackId>, item_type_id:1, purchase_type_id:1, audio_format_id:1, source_type_id:1, country_id:3, cart:<cartId>}`
  → `201`. `item_type_id` 1=track, 2=release.
  - `item_type` as a **string** → `400 {"detail":"Item does not exist"}` (misleading —
    it reads like a bad track id, but the field name is wrong).
  - Omitting `source_type_id` → `400 {"source_type_id":["This field is required."]}`.
    It's write-only, so it never appears in a GET; `1` works.
- Already-owned tracks return `400 {"AlreadyPurchased":{…}}` — treat as SUCCESS-adjacent
  (Keith owns it, don't re-buy). Round Circle hit this.
- `OPTIONS` on the items endpoint is **403**, so DRF metadata can't be used to discover
  fields — copy the shape off the real items in a GET instead.

Matching lesson (cost a whole pass): score title similarity on the **title alone**,
never `title + mix_name` — appending "(Original Mix)" dilutes a length-aware
containment score below threshold and rejects perfect matches. And guard **version
words** (reborn/rework/remaster/remix/edit/flip/bootleg/vip/dub): the wanted side's
set must be a subset of the candidate's, or "Release Me REBORN" happily matches the
plain "Release Me (Original Mix)". See [[project_sc_match]], [[project_radio_release_pipeline]].
