---
name: reference-supabase-secrets-are-write-only
description: "The Supabase Management API returns a SHA-256 digest in a secret's value field, not the value — never judge a credential from a read"
metadata: 
  node_type: memory
  type: reference
  originSessionId: 8ebdc8d3-10da-4c31-9305-ba69f006c2f9
  modified: 2026-08-25T17:41:52.338Z
---

`GET /v1/projects/{ref}/secrets` returns a **SHA-256 digest** in each secret's
`value` field. Every secret reads back as 64 hex characters no matter what it
holds. Proven by setting `EBAY_DELETION_ENDPOINT` to a URL known exactly and
confirming the read-back equalled `sha256(url)`.

**Consequences:**
- A secret is **write-only**. It can be written and then *exercised*, never
  inspected.
- There is **no rename and no copy**. Moving a value to a different name needs
  plaintext nobody has — the user must re-enter it. (They *can* see it: the
  Supabase dashboard → Edge Functions → Secrets has a reveal icon.)
- **Never characterise a credential from a read.** Length, charset, prefix,
  hyphens — all properties of the digest, not the secret.

**Why this matters (2026-08-25):** Keith's eBay keys were saved under the wrong
names (`App ID` / `Cert ID`). Reading them back gave 64-hex, which I described as
"not an eBay keyset" and told him so — twice, against his explicit correction. I
then copied those *digests* into `EBAY_CLIENT_ID` / `EBAY_CLIENT_SECRET` and
tested the digests against eBay's OAuth endpoint; its `401 invalid_client` read as
confirmation. It was a hash of his password failing to be his password.

The real cause was what he'd found himself: eBay **disables a production keyset**
until Marketplace Account Deletion compliance is in place, and a disabled keyset
returns the same `401 invalid_client` as a wrong key.

**How to apply:** the only valid test of a credential is to send it to the system
that owns it. If an inspection contradicts what Keith says he entered, distrust
the inspection first. Related: [[feedback-pause-before-major-changes]] and the
same proxy-instead-of-invariant shape as LEARNINGS §37/§49/§51.

Also: Supabase **edge-function logs are empty on this plan** — `function_edge_logs`
returns zero rows for every function, including ones invoked minutes earlier.
Never read an empty result there as "it was never called"; run the unfiltered
control query first.
