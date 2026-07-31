// Supabase Edge Function: fan a push notification out to every subscribed
// device in a club. Deploy with:  supabase functions deploy push
// Secrets needed:  VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT (mailto:you@club.org)
import { createClient } from "npm:@supabase/supabase-js@2";
import webpush from "npm:web-push@3.6.7";

Deno.serve(async (req) => {
  try {
    const { clubId, title, body } = await req.json();
    if (!clubId || !title) return new Response("bad request", { status: 400 });

    // Verify the caller is a member of this club (uses the caller's JWT).
    const userClient = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } } },
    );
    const { data: isMember } = await userClient.rpc("is_club_member", { c: clubId });
    if (!isMember) return new Response("forbidden", { status: 403 });

    // Read subscriptions with the service role and send.
    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );
    webpush.setVapidDetails(
      Deno.env.get("VAPID_SUBJECT") ?? "mailto:club@example.org",
      Deno.env.get("VAPID_PUBLIC_KEY")!,
      Deno.env.get("VAPID_PRIVATE_KEY")!,
    );
    const { data: subs } = await admin
      .from("push_subscriptions").select("id, subscription").eq("club_id", clubId);

    const payload = JSON.stringify({ title, body });
    const results = await Promise.allSettled(
      (subs ?? []).map((s) => webpush.sendNotification(s.subscription, payload)),
    );
    // Prune dead subscriptions (endpoint gone = 404/410)
    const dead = (subs ?? []).filter((_, i) => {
      const r = results[i];
      return r.status === "rejected" && [404, 410].includes(r.reason?.statusCode);
    }).map((s) => s.id);
    if (dead.length) await admin.from("push_subscriptions").delete().in("id", dead);

    return new Response(JSON.stringify({ sent: results.filter((r) => r.status === "fulfilled").length }), {
      headers: { "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(String(e), { status: 500 });
  }
});
