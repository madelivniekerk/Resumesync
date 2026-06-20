// Supabase Edge Function — receives Pay Advanced webhooks and adds credits to user account
// Deploy: supabase functions deploy pay-advanced-webhook
// Webhook URL: https://svihqlnlfdavkmhjbqrd.supabase.co/functions/v1/pay-advanced-webhook
//
// Required Supabase secrets:
//   PAY_ADVANCED_SECRET  — webhook secret from Pay Advanced dashboard

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const PACK_CREDITS: Record<string, number> = {
  pack5:  5,
  pack20: 20,
  pack30: 30,
  pack50: 50,
};

// Amount-based fallback (in case paymentref is truncated or missing)
function creditsFromAmount(amount: number): number {
  if (amount >= 8   && amount < 12)  return 5;   // A$9
  if (amount >= 14  && amount < 17)  return 20;  // A$15
  if (amount >= 19  && amount < 22)  return 30;  // A$20
  if (amount >= 29  && amount < 33)  return 50;  // A$30
  return 0;
}

async function verifySignature(body: string, receivedSig: string, secret: string): Promise<boolean> {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sigBuffer = await crypto.subtle.sign("HMAC", key, encoder.encode(body));
  const base64Sig = btoa(String.fromCharCode(...new Uint8Array(sigBuffer)));
  return base64Sig === receivedSig;
}

Deno.serve(async (req: Request) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  const rawBody = await req.text();

  const secret = Deno.env.get("PAY_ADVANCED_SECRET") || "";
  if (secret) {
    const sig = req.headers.get("x-payadvantage-signature") || "";
    if (!sig) {
      console.error("Missing x-payadvantage-signature header");
      return new Response("Missing signature", { status: 401 });
    }
    const valid = await verifySignature(rawBody, sig, secret);
    if (!valid) {
      console.error("Signature mismatch — possible spoofed request");
      return new Response("Invalid signature", { status: 401 });
    }
  }

  let events: Array<Record<string, unknown>>;
  try {
    const parsed = JSON.parse(rawBody);
    events = Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return new Response("Invalid JSON", { status: 400 });
  }

  for (const event of events) {
    if (typeof event !== "object" || event === null || !event.Code || !event.Event) {
      return new Response("Invalid schema", { status: 400 });
    }
  }

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  for (const event of events) {
    const eventName = String(event.Event || "");
    console.log("Received event:", eventName);

    if (eventName === "webhook_endpoint.armed") {
      console.log("Webhook endpoint armed successfully");
      continue;
    }

    if (eventName !== "payment.created") continue;

    const data = event.Data as Record<string, unknown>;
    if (!data) continue;

    // paymentref format: email|pack5 / email|pack20 / email|pack30 / email|pack50
    const externalRef = String(data.ExternalID || data.ExternalReference || "");
    const amount = Number(data.Amount || 0);

    let email = "";
    let credits = 0;

    if (externalRef.includes("|")) {
      const [refEmail, refPack] = externalRef.split("|");
      email = refEmail.trim().toLowerCase();
      credits = PACK_CREDITS[refPack] ?? 0;
    }

    // Fall back to amount-based detection
    if (!credits) credits = creditsFromAmount(amount);

    if (!email || !credits) {
      console.error("Cannot identify user or pack — ExternalReference:", externalRef, "Amount:", amount);
      continue;
    }

    // Fetch current credits and increment
    const { data: rows, error: fetchErr } = await supabase
      .from("profiles")
      .select("credits")
      .eq("email", email)
      .single();

    if (fetchErr || !rows) {
      console.error("Could not fetch profile for", email, fetchErr);
      continue;
    }

    const newCredits = (Number(rows.credits) || 0) + credits;
    const { error: updateErr } = await supabase
      .from("profiles")
      .update({ credits: newCredits })
      .eq("email", email);

    if (updateErr) {
      console.error("Supabase update failed for", email, updateErr);
    } else {
      console.log(`Added ${credits} credits to ${email} — new balance: ${newCredits}`);
    }
  }

  return new Response(JSON.stringify({ ok: true }), {
    status: 202,
    headers: { "Content-Type": "application/json" },
  });
});
