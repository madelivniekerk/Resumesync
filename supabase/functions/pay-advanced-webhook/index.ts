// Supabase Edge Function — receives Pay Advanced webhooks and upgrades user tier
// Deploy: supabase functions deploy pay-advanced-webhook
// Webhook URL: https://svihqlnlfdavkmhjbqrd.supabase.co/functions/v1/pay-advanced-webhook
//
// Required Supabase secrets:
//   PAY_ADVANCED_SECRET  — webhook secret from Pay Advanced dashboard

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// Constant-time string comparison to prevent timing attacks
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
  const hexSig = Array.from(new Uint8Array(sigBuffer))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");

  // Pay Advanced may send as hex or with a prefix — try both
  const clean = receivedSig.replace(/^sha256=/, "");
  if (hexSig.length !== clean.length) return false;

  let diff = 0;
  for (let i = 0; i < hexSig.length; i++) {
    diff |= hexSig.charCodeAt(i) ^ clean.charCodeAt(i);
  }
  return diff === 0;
}

Deno.serve(async (req: Request) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  const rawBody = await req.text();

  // Verify signature if secret is configured
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

  // Pay Advanced sends an array of events
  let events: Array<Record<string, unknown>>;
  try {
    const parsed = JSON.parse(rawBody);
    events = Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return new Response("Invalid JSON", { status: 400 });
  }

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  for (const event of events) {
    const eventName = String(event.Event || "");
    console.log("Received event:", eventName);

    // Arming handshake — Pay Advanced sends this to verify the endpoint is live
    if (eventName === "webhook_endpoint.armed") {
      console.log("Webhook endpoint armed successfully");
      continue;
    }

    // Only process successful payments
    if (eventName !== "payment.created") continue;

    const data = event.Data as Record<string, unknown>;
    if (!data) continue;

    // ExternalReference is set by us when building the checkout URL as "email|plan"
    const externalRef = String(data.ExternalReference || "");
    const amount = Number(data.Amount || 0);

    let email = "";
    let tier = "";

    if (externalRef.includes("|")) {
      const [refEmail, refPlan] = externalRef.split("|");
      email = refEmail.trim().toLowerCase();
      if (refPlan === "starter" || refPlan === "unlimited") tier = refPlan;
    }

    // Fall back to amount-based tier detection
    if (!tier) {
      if (amount >= 9 && amount < 12)   tier = "starter";
      if (amount >= 14 && amount < 17)  tier = "unlimited";
    }

    if (!email || !tier) {
      console.error("Cannot identify user or tier — ExternalReference:", externalRef, "Amount:", amount);
      // Return 202 so Pay Advanced doesn't keep retrying
      continue;
    }

    const now   = new Date();
    const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

    const { error } = await supabase
      .from("profiles")
      .update({ tier, analyses_used: 0, analyses_month: month })
      .eq("email", email);

    if (error) {
      console.error("Supabase update failed for", email, error);
    } else {
      console.log(`Upgraded ${email} → ${tier}`);
    }
  }

  // Pay Advanced requires 202 Accepted
  return new Response(JSON.stringify({ ok: true }), {
    status: 202,
    headers: { "Content-Type": "application/json" },
  });
});
