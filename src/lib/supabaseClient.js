import { createClient } from "@supabase/supabase-js";

const supabaseUrl = typeof import.meta !== "undefined" ? String(import.meta.env?.VITE_SUPABASE_URL || "").trim() : "";
const supabaseAnonKey = typeof import.meta !== "undefined" ? String(import.meta.env?.VITE_SUPABASE_ANON_KEY || "").trim() : "";
const workspaceId = typeof import.meta !== "undefined" ? String(import.meta.env?.VITE_SUPABASE_WORKSPACE_ID || "main-workspace").trim() : "main-workspace";

export const supabaseEnabled = Boolean(supabaseUrl && supabaseAnonKey);
export const supabaseWorkspaceId = workspaceId || "main-workspace";

export const supabase = supabaseEnabled
  ? createClient(supabaseUrl, supabaseAnonKey, {
      auth: {
        persistSession: true,
        autoRefreshToken: true,
      },
    })
  : null;

