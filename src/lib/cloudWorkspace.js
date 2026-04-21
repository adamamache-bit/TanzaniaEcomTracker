import { supabase, supabaseEnabled, supabaseWorkspaceId } from "./supabaseClient";

export async function getCloudSession() {
  if (!supabaseEnabled || !supabase) return { session: null, user: null };
  const { data, error } = await supabase.auth.getSession();
  if (error) throw error;
  return {
    session: data.session || null,
    user: data.session?.user || null,
  };
}

export function onCloudAuthStateChange(callback) {
  if (!supabaseEnabled || !supabase) return () => {};
  const { data } = supabase.auth.onAuthStateChange((_event, session) => {
    callback({
      session: session || null,
      user: session?.user || null,
    });
  });
  return () => data.subscription.unsubscribe();
}

export async function signInCloud({ email, password }) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data;
}

export async function signUpCloud({ email, password }) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  const { data, error } = await supabase.auth.signUp({ email, password });
  if (error) throw error;
  return data;
}

export async function signOutCloud() {
  if (!supabaseEnabled || !supabase) return;
  const { error } = await supabase.auth.signOut();
  if (error) throw error;
}

export async function loadCloudWorkspace(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");

  const { data, error } = await supabase
    .from("workspaces")
    .select("id,name,version,updated_at,state")
    .eq("id", workspaceId)
    .maybeSingle();

  if (error) throw error;

  if (!data) {
    return {
      id: workspaceId,
      version: 0,
      updatedAt: null,
      state: null,
    };
  }

  return {
    id: data.id,
    version: Number(data.version || 0),
    updatedAt: data.updated_at || null,
    state: data.state || null,
    name: data.name || null,
  };
}

export async function saveCloudWorkspace(state, { workspaceId = supabaseWorkspaceId, userId = null } = {}) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");

  const current = await loadCloudWorkspace(workspaceId);
  const nextVersion = Math.max(Number(current.version || 0) + 1, Date.now());
  const payload = {
    id: workspaceId,
    name: "Main Workspace",
    version: nextVersion,
    updated_at: new Date().toISOString(),
    updated_by: userId,
    state,
  };

  const { data, error } = await supabase.from("workspaces").upsert(payload).select("id,version,updated_at").single();
  if (error) throw error;

  return {
    id: data.id,
    version: Number(data.version || 0),
    updatedAt: data.updated_at || null,
  };
}

export function subscribeToCloudWorkspace(workspaceId = supabaseWorkspaceId, onChange) {
  if (!supabaseEnabled || !supabase) return () => {};

  const channel = supabase
    .channel(`workspace:${workspaceId}`)
    .on(
      "postgres_changes",
      {
        event: "*",
        schema: "public",
        table: "workspaces",
        filter: `id=eq.${workspaceId}`,
      },
      (payload) => {
        const next = payload?.new || null;
        if (!next) return;
        onChange({
          id: next.id,
          version: Number(next.version || 0),
          updatedAt: next.updated_at || null,
          state: next.state || null,
        });
      }
    )
    .subscribe();

  return () => {
    supabase.removeChannel(channel);
  };
}
