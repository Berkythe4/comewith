// ingest-email
//
// "Create things by email from the road." You email a template to an inbound
// address (wired via an email provider's inbound webhook — see SETUP below);
// that provider POSTs the message here; this parses it and creates the record.
//
// SECURITY: an email's `from` is trivially spoofed, so the real gate is a shared
// SECRET that must arrive either as ?key=<secret> on the webhook URL OR as a
// "Key: <secret>" line in the body. No valid secret → 401, nothing created.
// Secret: INGEST_SECRET (project secret). Optional: RESEND_API_KEY to email a
// confirmation back to the sender.
//
// TEMPLATE (subject sets the type; body is `Field: value` lines):
//   Subject:  EVENT: Summer Warehouse Party
//   Body:
//     Key: <the secret>
//     Date: 2026-09-15
//     Venue: Brooklyn Storehouse
//     Series: Come With Parties
//     Capacity: 300
//     Doors: 22:00
//     Ticket URL: https://...
//     Description: Peak-summer rave.
//
// Also supports  EXPENSE:  (Amount, Category, Date, Vendor, Description)
//           and  TASK:     (Due, Priority, Description)
//
// SETUP (one-time, external): point an inbound route at
//   https://<ref>.functions.supabase.co/ingest-email?key=<INGEST_SECRET>
// e.g. a Cloudflare Email Routing worker, Resend inbound, or SendGrid Inbound
// Parse forwarding mail sent to (say) inbox@comewith.org.
//
// Body accepted: Resend/SendGrid-style JSON, or a generic {from, subject, text}.

import { createClient } from "npm:@supabase/supabase-js@2";

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "authorization, content-type", "Access-Control-Allow-Methods": "POST, OPTIONS" };
const JH = { ...CORS, "Content-Type": "application/json" };
const ok = (o: unknown) => new Response(JSON.stringify(o), { headers: JH });
const bad = (s: number, m: string) => new Response(JSON.stringify({ error: m }), { status: s, headers: JH });

const SERIES = ["Come With Parties", "Dance Infusion", "Come With Production", "Content Creation", "Bookings"];
const slugify = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);

// Pull {from, subject, text} out of whatever shape the inbound provider sent.
function normalize(b: any): { from: string; subject: string; text: string } {
  const from = (b.from?.email || b.from?.address || b.from || b.sender || b.envelope?.from || "").toString();
  const subject = (b.subject || b.Subject || "").toString();
  const text = (b.text || b["body-plain"] || b.plain || b.TextBody || b.body || b.html || "").toString();
  return { from: from.replace(/.*<([^>]+)>.*/, "$1").trim(), subject: subject.trim(), text };
}
// Parse "Field: value" lines (case-insensitive keys). Multi-word values kept.
function fields(text: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const raw of text.split(/\r?\n/)) {
    const m = raw.match(/^\s*([A-Za-z][A-Za-z _/]{0,30}?)\s*[:\-]\s*(.+?)\s*$/);
    if (m) out[m[1].toLowerCase().replace(/\s+/g, "_")] = m[2].trim();
  }
  return out;
}
const parseDate = (s?: string) => {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(+d) ? null : d.toISOString().slice(0, 10);
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return bad(405, "POST only");

  const SECRET = Deno.env.get("INGEST_SECRET");
  if (!SECRET) return bad(500, "INGEST_SECRET not configured");

  const url = new URL(req.url);
  const b = await req.json().catch(() => ({}));
  const { from, subject, text } = normalize(b);
  const f = fields(text);

  // Gate: secret from the URL OR a Key: line. Constant-ish compare.
  const provided = url.searchParams.get("key") || f["key"] || f["token"] || "";
  if (provided !== SECRET) return bad(401, "bad or missing key");

  const admin = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!);
  const m = subject.match(/^\s*(EVENT|EXPENSE|TASK)\s*:\s*(.+)$/i);
  if (!m) return ok({ ignored: true, reason: "Subject must start with EVENT: / EXPENSE: / TASK:" });
  const type = m[1].toUpperCase();
  const titleText = m[2].trim();

  let created: any = null, summary = "";
  try {
    if (type === "EVENT") {
      const series = SERIES.find((s) => s.toLowerCase() === (f["series"] || "").toLowerCase()) || null;
      const date = parseDate(f["date"]) || parseDate(f["event_date"]);
      if (!date) return ok({ error: "EVENT needs a Date: line (YYYY-MM-DD)" });
      let venue_id: string | null = null;
      if (f["venue"]) {
        const { data: v } = await admin.from("venues").select("id").ilike("name", f["venue"]).is("deleted_at", null).limit(1).maybeSingle();
        venue_id = v?.id || null;
      }
      const row = {
        name: titleText, slug: slugify(titleText) + "-" + date, series, event_date: date,
        status: "planning", stage: "planning",
        venue_id, capacity: f["capacity"] ? Number(f["capacity"].replace(/\D/g, "")) || null : null,
        doors_time: f["doors"] || f["doors_time"] || null, ticket_url: f["ticket_url"] || f["tickets"] || null,
        description: f["description"] || f["notes"] || null,
      };
      const { data, error } = await admin.from("events").insert(row).select("id,name,event_date").single();
      if (error) throw error;
      created = data; summary = `Event "${data.name}" on ${data.event_date}${series ? " · " + series : ""}${venue_id ? "" : (f["venue"] ? " (venue not matched)" : "")}`;
    } else if (type === "EXPENSE") {
      const amount = Number((f["amount"] || "").replace(/[^0-9.]/g, ""));
      if (!amount) return ok({ error: "EXPENSE needs an Amount: line" });
      const row = { description: titleText, amount, category: f["category"] || null, vendor: f["vendor"] || null, date: parseDate(f["date"]) || new Date().toISOString().slice(0, 10) };
      const { data, error } = await admin.from("expenses").insert(row).select("id,amount,description").single();
      if (error) throw error;
      created = data; summary = `Expense $${data.amount} — ${data.description}`;
    } else if (type === "TASK") {
      const row = { title: titleText, status: "todo", source: "manual", description: f["description"] || null, due_date: parseDate(f["due"]) || parseDate(f["due_date"]), priority: (f["priority"] || "").toLowerCase() || null };
      const { data, error } = await admin.from("tasks").insert(row).select("id,title").single();
      if (error) throw error;
      created = data; summary = `Task "${data.title}"`;
    }
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    await confirm(from, `✗ Couldn't create ${type}`, `Come With ingest hit an error:\n\n${msg}\n\nDouble-check the template and try again.`);
    return ok({ error: msg });
  }

  await confirm(from, `✓ Created: ${summary}`, `Come With ingested your email and created:\n\n${summary}\n\nOpen the dashboard to add details.`);
  return ok({ success: true, type, created, summary });
});

// Best-effort confirmation back to the sender (only reached after the secret
// check passed, so this can't be used to spam arbitrary addresses).
async function confirm(to: string, subject: string, body: string) {
  const key = Deno.env.get("RESEND_API_KEY");
  if (!key || !to || !/@/.test(to)) return;
  try {
    await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Authorization": "Bearer " + key, "Content-Type": "application/json" },
      body: JSON.stringify({ from: Deno.env.get("FROM_EMAIL") || "Come With <berky@comewith.org>", to, subject, text: body }),
    });
  } catch { /* ignore */ }
}
