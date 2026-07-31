import { supabase, supabaseConfigured } from "./supabase";

let localVersion = 0;

export function storageEnabled() { return supabaseConfigured; }

export async function getClubDoc(clubId) {
  if (!supabaseConfigured) {
    const v = localStorage.getItem("club:" + clubId);
    return v ? JSON.parse(v) : null;
  }
  const { data, error } = await supabase
    .from("club_data").select("doc,version").eq("club_id", clubId).single();
  if (error) throw error;
  localVersion = data.version;
  return data.doc && Object.keys(data.doc).length ? data.doc : null;
}

/** Optimistic-concurrency save. Throws { fresh } on version conflict. */
export async function saveClubDoc(clubId, doc) {
  if (!supabaseConfigured) {
    localStorage.setItem("club:" + clubId, JSON.stringify(doc));
    return;
  }
  const { data, error } = await supabase.rpc("save_club_doc", {
    p_club: clubId, p_doc: doc, p_expected: localVersion,
  });
  if (error) throw error;
  if (data === null) {
    const fresh = await getClubDoc(clubId); // also refreshes localVersion
    const err = new Error("conflict");
    err.fresh = fresh;
    throw err;
  }
  localVersion = data;
}

/** Live sync: refetch the doc whenever another member saves. Returns unsubscribe fn. */
export function subscribeClubDoc(clubId, onDoc) {
  if (!supabaseConfigured) return () => {};
  const ch = supabase
    .channel("club_data:" + clubId)
    .on("postgres_changes",
      { event: "UPDATE", schema: "public", table: "club_data", filter: `club_id=eq.${clubId}` },
      async (payload) => {
        const v = payload.new?.version ?? 0;
        if (v > localVersion) {
          try { const doc = await getClubDoc(clubId); if (doc) onDoc(doc); } catch (e) {}
        }
      })
    .subscribe();
  return () => supabase.removeChannel(ch);
}
