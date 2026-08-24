import { supabase, supabaseConfigured } from "./supabase";

/** Manually trigger dues-invoice emails for charges due today in this club. */
export async function sendInvoicesNow(clubId) {
  if (!supabaseConfigured) throw new Error("Not available in demo mode");
  const { data, error } = await supabase.functions.invoke("invoices", { body: { clubId } });
  if (error) throw error;
  return data;
}
