import { supabase, supabaseConfigured } from "./supabase";

export function filesEnabled() { return supabaseConfigured; }

/** Upload to the club-files bucket; returns a public URL. */
export async function uploadFile(clubId, file) {
  const safe = file.name.replace(/[^\w.\-]+/g, "_");
  const path = `${clubId}/${Date.now()}-${Math.random().toString(36).slice(2, 8)}-${safe}`;
  const { error } = await supabase.storage.from("club-files")
    .upload(path, file, { upsert: false, contentType: file.type || "application/octet-stream" });
  if (error) throw error;
  const { data } = supabase.storage.from("club-files").getPublicUrl(path);
  return data.publicUrl;
}
