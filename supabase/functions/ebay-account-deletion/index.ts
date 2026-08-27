// eBay Marketplace Account Deletion / Closure notification endpoint.
//
// WHY THIS EXISTS. eBay will not enable a PRODUCTION keyset until the account
// holds a working notification endpoint (or an exemption). Until then every
// OAuth call answers 401 invalid_client — which is exactly what Gear Watch was
// getting, and it reads identically to a wrong password. The keys were right
// the whole time; the keyset was disabled.
//
// WHAT eBay REQUIRES, and the two ways this is got wrong:
//
//   1. A GET challenge. eBay calls this URL with ?challenge_code=… and expects
//      200 + {"challengeResponse": sha256(challengeCode + verificationToken +
//      endpoint)} as HEX. The three parts are concatenated in THAT ORDER with no
//      separator, and `endpoint` must be the URL byte-for-byte as registered —
//      a trailing slash, http vs https, or a stray query string all produce a
//      valid-looking hash that eBay rejects. It is therefore read from a secret
//      rather than reconstructed from the request, because behind a proxy
//      req.url is not necessarily the URL eBay was told about.
//
//   2. A POST notification, when a real user deletes their eBay account. Ack it
//      with 200 promptly; eBay retries and then disables the keyset if it does
//      not get one.
//
// WHAT WE ACTUALLY DO WITH A DELETION. Nothing, and that is the honest answer:
// Gear Watch reads public marketplace LISTINGS to look for stolen equipment. It
// stores listing id, title, price, location and seller username — no eBay user
// accounts, no buyer data, nothing keyed to an eBay user id. There is no
// personal record to erase. The notification is logged so compliance can be
// evidenced, and acknowledged.
//
// This endpoint MUST be deployed --no-verify-jwt: eBay calls it unauthenticated,
// and a 401 at the gateway looks to eBay like a dead endpoint.

const TOKEN = Deno.env.get("EBAY_VERIFICATION_TOKEN") || "";
const ENDPOINT = Deno.env.get("EBAY_DELETION_ENDPOINT") || "";

const J = { "Content-Type": "application/json" };

async function sha256Hex(s: string): Promise<string> {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

Deno.serve(async (req: Request) => {
  const url = new URL(req.url);

  // ── 1. the verification challenge ────────────────────────────────────────
  const challenge = url.searchParams.get("challenge_code");
  if (req.method === "GET" && challenge) {
    // A misconfiguration must NOT answer with a plausible-looking hash. eBay
    // would mark the endpoint verified and the keyset would still be dead, with
    // nothing anywhere saying why.
    if (!TOKEN || !ENDPOINT) {
      console.error("ebay-account-deletion: EBAY_VERIFICATION_TOKEN / EBAY_DELETION_ENDPOINT not set");
      return new Response(
        JSON.stringify({ error: "endpoint not configured — set EBAY_VERIFICATION_TOKEN and EBAY_DELETION_ENDPOINT" }),
        { status: 500, headers: J },
      );
    }
    const challengeResponse = await sha256Hex(challenge + TOKEN + ENDPOINT);
    return new Response(JSON.stringify({ challengeResponse }), { status: 200, headers: J });
  }

  // ── 2. an actual account-deletion notification ───────────────────────────
  if (req.method === "POST") {
    const body = await req.json().catch(() => ({}));
    const n = body?.notification?.data ?? {};
    // Log rather than store: there is no eBay user data here to delete, and
    // inventing a table to record other people's deletion requests in would
    // itself be holding data we were told to stop holding.
    console.log("ebay account deletion notification", JSON.stringify({
      notificationId: body?.notification?.notificationId ?? null,
      eventDate: body?.notification?.eventDate ?? null,
      username: n?.username ?? null,
      userId: n?.userId ?? null,
      action: "acknowledged — Gear Watch stores public listings only, no eBay user records",
    }));
    // 200 with no body is what eBay wants; anything slower or noisier just
    // increases the chance of a retry storm.
    return new Response(null, { status: 200 });
  }

  // A bare GET is eBay (or you) checking the endpoint is alive.
  if (req.method === "GET") {
    return new Response(JSON.stringify({
      ok: true,
      endpoint: "eBay marketplace account deletion",
      configured: !!(TOKEN && ENDPOINT),
    }), { status: 200, headers: J });
  }

  return new Response(JSON.stringify({ error: "method not allowed" }), { status: 405, headers: J });
});
