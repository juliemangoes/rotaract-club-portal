import { supabase, supabaseConfigured } from "./supabase";

function urlB64ToUint8Array(s) {
  const pad = "=".repeat((4 - (s.length % 4)) % 4);
  const b = (s + pad).replace(/-/g, "+").replace(/_/g, "/");
  const raw = atob(b);
  return Uint8Array.from([...raw].map((c) => c.charCodeAt(0)));
}

/** Ask permission, subscribe this device, and store the subscription for this club. */
export async function registerPush(clubId) {
  if (!supabaseConfigured) return;
  if (!("serviceWorker" in navigator) || !("PushManager" in window)) return;
  const vapid = import.meta.env.VITE_VAPID_PUBLIC_KEY;
  if (!vapid) return;
  const perm = await Notification.requestPermission();
  if (perm !== "granted") return;
  const reg = await navigator.serviceWorker.ready;
  const sub = await reg.pushManager.subscribe({
    userVisibleOnly: true,
    applicationServerKey: urlB64ToUint8Array(vapid),
  });
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return;
  await supabase.from("push_subscriptions").upsert(
    { user_id: user.id, club_id: clubId, subscription: sub.toJSON() },
    { onConflict: "user_id,club_id" }
  );
}

/** Fan a notification out to every subscribed device in the club (fire-and-forget). */
export async function sendPush(clubId, payload) {
  if (!supabaseConfigured) return;
  try { await supabase.functions.invoke("push", { body: { clubId, ...payload } }); } catch (e) {}
}
