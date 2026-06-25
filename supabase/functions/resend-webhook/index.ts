// resend-webhook
//
// Public endpoint that Resend POSTs to with delivery events
// (email.delivered / email.bounced / email.complained / email.opened /
// email.clicked / etc).
//
// Optional but recommended: RESEND_WEBHOOK_SECRET (svix-style signing).
// If set, verifies the svix signature on each request.
//
// For each event:
//   1. Insert a mailing_events row with the event_type, looking up
//      the campaign + subscriber via resend_event_id from the original
//      'sent' event (Resend includes the email_id in its payload).
//   2. For bounced/complained: flip subscribers.status to match so
//      send-campaign skips them on the next send.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const JSON_HEADERS = { ...CORS_HEADERS, "Content-Type": "application/json" };

function jsonError(s: number, m: string) {
  return new Response(JSON.stringify({ error: m }), { status: s, headers: JSON_HEADERS });
}

async function verifySvix(secret: string, rawBody: string, headers: Headers): Promise<boolean> {
  const svixId = headers.get("svix-id");
  const svixTs = headers.get("svix-timestamp");
  const svixSig = headers.get("svix-signature");
  if (!svixId || !svixTs || !svixSig) return false;

  // svix secret format: whsec_BASE64STRING
  const base64 = secret.startsWith("whsec_") ? secret.slice(6) : secret;
  const keyBytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const key = await crypto.subtle.importKey(
    "raw", keyBytes, { name: "HMAC", hash: "SHA-256" }, false, ["sign"],
  );
  const payload = `${svixId}.${svixTs}.${rawBody}`;
  const sigBuf = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload));
  const expected = btoa(String.fromCharCode(...new Uint8Array(sigBuf)));

  // svix-signature is a space-separated list of "v1,SIG" pairs
  return svixSig.split(" ").some((part) => {
    const [, sig] = part.split(",");
    return sig === expected;
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return jsonError(405, "POST only");

  try {
    const rawBody = await req.text();

    const secret = Deno.env.get("RESEND_WEBHOOK_SECRET");
    if (secret) {
      const ok = await verifySvix(secret, rawBody, req.headers);
      if (!ok) return jsonError(401, "invalid signature");
    }
    // If no secret set, accept all (Phase 11 should set the secret).

    const payload = JSON.parse(rawBody);
    // Resend payload shape:
    // { type: "email.delivered", created_at: "...", data: { email_id, to, from, subject, ... } }
    const eventType: string = (payload.type || "").replace(/^email\./, "");
    const resendEmailId: string | undefined = payload.data?.email_id;

    if (!eventType || !resendEmailId) {
      return jsonError(400, "missing type or email_id in payload");
    }

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    // Find the original 'sent' event to attribute campaign + subscriber.
    const { data: originalSent } = await admin
      .from("mailing_events")
      .select("campaign_id, subscriber_id")
      .eq("resend_event_id", resendEmailId)
      .eq("event_type", "sent")
      .maybeSingle();

    const newEventRow = {
      campaign_id: originalSent?.campaign_id || null,
      subscriber_id: originalSent?.subscriber_id || null,
      event_type: eventType,
      resend_event_id: resendEmailId,
      metadata: payload.data,
      occurred_at: payload.created_at || new Date().toISOString(),
    };

    await admin.from("mailing_events").insert(newEventRow);

    // For bounce/complaint, sync subscriber status.
    if ((eventType === "bounced" || eventType === "complained") && originalSent?.subscriber_id) {
      await admin
        .from("subscribers")
        .update({ status: eventType })
        .eq("id", originalSent.subscriber_id);
    }

    // --- Also correlate to a Conversation message (resend_id == email_id) so
    //     delivery + bounce status shows in the thread for the whole team. ---
    const { data: cmsg } = await admin
      .from("conversation_messages")
      .select("id, conversation_id")
      .eq("resend_id", resendEmailId)
      .maybeSingle();
    let conv_attributed = false;
    if (cmsg) {
      conv_attributed = true;
      await admin.from("conversation_messages").update({ status: eventType }).eq("id", cmsg.id);
      if (eventType === "bounced" || eventType === "complained" || eventType === "delivery_delayed") {
        const note = eventType === "bounced"
          ? "⚠ Delivery failed — the email bounced (the address rejected it). They did NOT receive it."
          : eventType === "complained"
          ? "⚠ Recipient marked this email as spam."
          : "⏳ Delivery delayed — still trying to reach the recipient.";
        await admin.from("conversation_messages").insert({
          conversation_id: cmsg.conversation_id, direction: "event", body: note,
          status: eventType, meta: payload.data,
        });
        await admin.from("conversations")
          .update({ last_message_at: new Date().toISOString() })
          .eq("id", cmsg.conversation_id);
      }
    }

    return new Response(
      JSON.stringify({ success: true, event_type: eventType, attributed: !!originalSent, conv_attributed }),
      { headers: JSON_HEADERS },
    );
  } catch (e) {
    return jsonError(500, "Unexpected: " + (e instanceof Error ? e.message : String(e)));
  }
});
