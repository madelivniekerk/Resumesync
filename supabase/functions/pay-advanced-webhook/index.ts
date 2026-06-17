// Supabase Edge Function — receives Pay Advanced webhook and upgrades user tier
// Deploy: supabase functions deploy pay-advanced-webhook
// Webhook URL (set in Pay Advanced dashboard):
//   https://svihqlnlfdavkmhjbqrd.supabase.co/functions/v1/pay-advanced-webhook

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const TIER_BY_AMOUNT: Record<string, string> = {
  "9.90":  "starter",
  "9.9":   "starter",
  "14.90": "unlimited",
  "14.9":  "unlimited",
};

const TIER_BY_REFERENCE: Record<string, string> = {
  "starter":   "starter",
  "unlimited": "unlimited",
};

Deno.serve(async (req: Request) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  let body: Record<string, unknown> = {};
  try {
    body = await req.json();
  } catch {
    return new Response("Invalid JSON", { status: 400 });
  }

  // Extract email — Pay Advanced typically sends it as 'email' or 'customer_email'
  const email =
    (body.email as string) ||
    (body.customer_email as string) ||
    (body.payer_email as string) ||
    "";

  if (!email) {
    console.error("Webhook missing email field", JSON.stringify(body));
    return new Response("Missing email", { status: 400 });
  }

  // Determine tier: prefer reference field, fall back to amount
  const reference = String(body.reference || body.plan || "").toLowerCase();
  const amount    = String(body.amount || body.total || body.payment_amount || "").replace(/[^0-9.]/g, "");

  const tier =
    TIER_BY_REFERENCE[reference] ||
    TIER_BY_AMOUNT[amount] ||
    null;

  if (!tier) {
    console.error("Cannot determine tier from webhook", JSON.stringify(body));
    // Return 200 so Pay Advanced doesn't retry — log for manual review
    return new Response(JSON.stringify({ ok: false, reason: "unknown tier" }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  // Only activate on successful payment status
  const status = String(body.status || body.payment_status || "success").toLowerCase();
  if (!["success", "paid", "completed", "approved"].includes(status)) {
    return new Response(JSON.stringify({ ok: false, reason: "payment not successful" }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }

  // Update profiles table
  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
  );

  const now   = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

  const { error } = await supabase
    .from("profiles")
    .update({ tier, analyses_used: 0, analyses_month: month })
    .eq("email", email);

  if (error) {
    console.error("Supabase update failed", error);
    return new Response(JSON.stringify({ ok: false, reason: error.message }), {
      status: 500,
      headers: { "Content-Type": "application/json" },
    });
  }

  console.log(`Upgraded ${email} to ${tier}`);
  return new Response(JSON.stringify({ ok: true, email, tier }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
});
