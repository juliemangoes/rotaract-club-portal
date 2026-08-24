// Supabase Edge Function: email a dues invoice to every member whose charge
// becomes due today. Deploy with:  supabase functions deploy invoices
// Secrets needed:  RESEND_API_KEY, INVOICE_FROM_EMAIL (optional, defaults to
// onboarding@resend.dev — Resend's shared sending domain, works with no setup)
//
// Two ways to call it:
//   1. No body / no clubId — scans every club for charges due today. Only the
//      service role (a scheduled Cron job) may do this.
//   2. { clubId } — scans just that club, for a manual "Send invoices now"
//      button. Caller must be a member of that club (their own JWT is checked).
import { createClient } from "npm:@supabase/supabase-js@2";

const ACTIVE_LIKE = ["Active", "On Leave"];
const CHARGE_LABEL: Record<string, string> = {
  monthly: "Monthly dues", district: "District dues", ri: "Rotary International dues",
  penalty: "Penalty", late: "Late charge", carryforward: "Balance carried forward", other: "Charge",
};
const money = (n: number, c = "$") => `${c}${Number(n || 0).toFixed(2)}`;
const todayStr = () => new Date().toISOString().slice(0, 10);

async function sendEmail(to: string, subject: string, html: string) {
  const apiKey = Deno.env.get("RESEND_API_KEY");
  if (!apiKey) throw new Error("RESEND_API_KEY not configured");
  const from = Deno.env.get("INVOICE_FROM_EMAIL") || "onboarding@resend.dev";
  const res = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({ from, to, subject, html }),
  });
  if (!res.ok) throw new Error(`Resend ${res.status}: ${await res.text()}`);
}

function invoiceHtml(clubName: string, memberName: string, label: string, amount: number, currency: string, dueDate: string) {
  return `<div style="font-family:system-ui,sans-serif;color:#2A1420;max-width:480px;margin:0 auto">
    <h2 style="color:#D41367">${clubName}</h2>
    <p>Hi ${memberName},</p>
    <p>This is a reminder that the following is due today, ${dueDate}:</p>
    <table style="width:100%;border-collapse:collapse;margin:16px 0">
      <tr><td style="padding:8px 0;border-bottom:1px solid #E9E1E6">${label}</td>
          <td style="padding:8px 0;border-bottom:1px solid #E9E1E6;text-align:right;font-weight:700">${money(amount, currency)}</td></tr>
    </table>
    <p>Please arrange payment with your club treasurer at your earliest convenience.</p>
    <p style="color:#8A7580;font-size:12px">This is an automated reminder from your club's portal.</p>
  </div>`;
}

// Sends invoices for charges due today within one already-loaded club doc.
// Mutates `doc` in place (marks charges as invoiced) and returns how many were sent.
async function processClub(clubName: string, doc: any): Promise<number> {
  const year = doc.years?.find((y: any) => y.id === doc.activeYearId);
  if (!year || !Array.isArray(doc.charges)) return 0;
  const today = todayStr();
  let sent = 0;
  for (const charge of doc.charges) {
    if (charge.reversed || charge.invoicedAt) continue;
    if (charge.yearId !== year.id || charge.dueDate !== today) continue;
    const member = doc.members?.find((m: any) => m.id === charge.memberId);
    if (!member?.email || !ACTIVE_LIKE.includes(member.status)) continue; // never invoice exempt statuses
    try {
      await sendEmail(
        member.email,
        `Dues due today: ${charge.label || CHARGE_LABEL[charge.kind] || "Charge"}`,
        invoiceHtml(clubName, member.name, charge.label || CHARGE_LABEL[charge.kind] || "Charge", charge.amount, doc.duesConfig?.currency || "$", charge.dueDate),
      );
      charge.invoicedAt = new Date().toISOString();
      sent++;
    } catch (e) {
      console.error(`Failed to email ${member.email}:`, e);
    }
  }
  return sent;
}

Deno.serve(async (req) => {
  try {
    let clubId: string | undefined;
    try { ({ clubId } = await req.json()); } catch { /* empty body is fine */ }

    const admin = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
    );

    if (clubId) {
      // Manual trigger: caller must be a member of this specific club.
      const userClient = createClient(
        Deno.env.get("SUPABASE_URL")!,
        Deno.env.get("SUPABASE_ANON_KEY")!,
        { global: { headers: { Authorization: req.headers.get("Authorization") ?? "" } } },
      );
      const { data: isMember } = await userClient.rpc("is_club_member", { c: clubId });
      if (!isMember) return new Response("forbidden", { status: 403 });
    } else {
      // Full scan across all clubs: only the service role (scheduled Cron) may do this.
      const auth = req.headers.get("Authorization") ?? "";
      if (auth !== `Bearer ${Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")}`) return new Response("forbidden", { status: 403 });
    }

    let query = admin.from("club_data").select("club_id, doc, version, clubs(name)");
    if (clubId) query = query.eq("club_id", clubId);
    const { data: rows, error } = await query;
    if (error) throw error;

    let totalSent = 0;
    for (const row of rows ?? []) {
      const doc = row.doc;
      const sent = await processClub((row as any).clubs?.name || "Your club", doc);
      if (sent > 0) {
        totalSent += sent;
        await admin.from("club_data").update({ doc, version: row.version + 1 }).eq("club_id", row.club_id).eq("version", row.version);
      }
    }

    return new Response(JSON.stringify({ invoicesSent: totalSent }), { headers: { "Content-Type": "application/json" } });
  } catch (e) {
    return new Response(String(e), { status: 500 });
  }
});
