import { supabase, supabaseEnabled, supabaseWorkspaceId } from "./supabaseClient";
import { hasMeaningfulWorkspaceData } from "./appLogic";

function isWorkspaceBackupsTableMissing(error) {
  const message = String(error?.message || error?.details || error?.hint || "").toLowerCase();
  return error?.code === "42P01" || (message.includes("workspace_backups") && (message.includes("does not exist") || message.includes("not found")));
}

function buildWorkspaceBackupSummary(state = {}) {
  return {
    products: Array.isArray(state?.products) ? state.products.length : 0,
    tracking: Array.isArray(state?.tracking) ? state.tracking.length : 0,
    customers: Array.isArray(state?.customers) ? state.customers.length : 0,
  };
}

async function ensureCloudWorkspaceMembership(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");

  const {
    data: { user },
    error: userError,
  } = await supabase.auth.getUser();

  if (userError) throw userError;
  if (!user) throw new Error("Cloud user session not found.");

  const { data: existingMembership, error: selectError } = await supabase
    .from("workspace_members")
    .select("workspace_id,user_id,role")
    .eq("workspace_id", workspaceId)
    .eq("user_id", user.id)
    .maybeSingle();

  if (selectError) throw selectError;
  if (existingMembership) return existingMembership;

  const { data: insertedMembership, error: insertError } = await supabase
    .from("workspace_members")
    .insert({
      workspace_id: workspaceId,
      user_id: user.id,
      role: "admin",
    })
    .select("workspace_id,user_id,role")
    .single();

  if (insertError) throw insertError;
  return insertedMembership;
}

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
  await ensureCloudWorkspaceMembership();
  return data;
}

export async function signUpCloud({ email, password }) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  const { data, error } = await supabase.auth.signUp({ email, password });
  if (error) throw error;
  await ensureCloudWorkspaceMembership();
  return data;
}

export async function signOutCloud() {
  if (!supabaseEnabled || !supabase) return;
  const { error } = await supabase.auth.signOut();
  if (error) throw error;
}

export async function loadCloudWorkspace(workspaceId = supabaseWorkspaceId) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  await ensureCloudWorkspaceMembership(workspaceId);

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

export async function listCloudWorkspaceBackups(workspaceId = supabaseWorkspaceId, limit = 12) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  await ensureCloudWorkspaceMembership(workspaceId);

  const { data, error } = await supabase
    .from("workspace_backups")
    .select("id,workspace_id,workspace_version,created_at,created_by,reason,summary")
    .eq("workspace_id", workspaceId)
    .order("created_at", { ascending: false })
    .limit(limit);

  if (error) {
    if (isWorkspaceBackupsTableMissing(error)) {
      return {
        available: false,
        items: [],
        notice: "Run the latest Supabase schema to enable cloud restore history.",
      };
    }
    throw error;
  }

  return {
    available: true,
    items: Array.isArray(data) ? data : [],
    notice: "",
  };
}

async function createCloudWorkspaceBackup(state, { workspaceId, userId = null, version = 0, reason = "autosave" } = {}) {
  if (!hasMeaningfulWorkspaceData(state)) {
    return {
      available: true,
      saved: false,
      entry: null,
      notice: "Skipping empty backup snapshot.",
    };
  }

  const latestResponse = await supabase
    .from("workspace_backups")
    .select("id,state,workspace_version,created_at")
    .eq("workspace_id", workspaceId)
    .order("created_at", { ascending: false })
    .limit(1)
    .maybeSingle();

  if (latestResponse.error) {
    if (isWorkspaceBackupsTableMissing(latestResponse.error)) {
      return {
        available: false,
        saved: false,
        entry: null,
        notice: "Cloud restore history table is missing.",
      };
    }
    throw latestResponse.error;
  }

  const latest = latestResponse.data || null;
  const serializedState = JSON.stringify(state || {});
  const latestSerializedState = latest?.state ? JSON.stringify(latest.state) : "";
  if (latest && latestSerializedState === serializedState) {
    return {
      available: true,
      saved: false,
      entry: latest,
      notice: "Latest backup already matches current state.",
    };
  }

  const summary = buildWorkspaceBackupSummary(state);
  const { data, error } = await supabase
    .from("workspace_backups")
    .insert({
      workspace_id: workspaceId,
      workspace_version: Number(version || 0),
      state,
      created_by: userId,
      reason,
      summary,
    })
    .select("id,workspace_id,workspace_version,created_at,created_by,reason,summary")
    .single();

  if (error) {
    if (isWorkspaceBackupsTableMissing(error)) {
      return {
        available: false,
        saved: false,
        entry: null,
        notice: "Cloud restore history table is missing.",
      };
    }
    throw error;
  }

  return {
    available: true,
    saved: true,
    entry: data || null,
    notice: "",
  };
}

export async function restoreCloudWorkspaceBackup(backupId, { workspaceId = supabaseWorkspaceId, userId = null } = {}) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  await ensureCloudWorkspaceMembership(workspaceId);

  const { data, error } = await supabase
    .from("workspace_backups")
    .select("id,workspace_id,state,workspace_version,created_at")
    .eq("workspace_id", workspaceId)
    .eq("id", backupId)
    .maybeSingle();

  if (error) {
    if (isWorkspaceBackupsTableMissing(error)) {
      throw new Error("Cloud restore history is not configured yet. Run the latest Supabase schema first.");
    }
    throw error;
  }

  if (!data?.state) {
    throw new Error("Backup snapshot not found.");
  }

  return saveCloudWorkspace(data.state, {
    workspaceId,
    userId,
    backupReason: `restore:${backupId}`,
  });
}

export async function saveCloudWorkspace(state, { workspaceId = supabaseWorkspaceId, userId = null, backupReason = "autosave" } = {}) {
  if (!supabaseEnabled || !supabase) throw new Error("Supabase is not configured.");
  await ensureCloudWorkspaceMembership(workspaceId);

  const current = await loadCloudWorkspace(workspaceId);
  if (!hasMeaningfulWorkspaceData(state) && hasMeaningfulWorkspaceData(current.state || {})) {
    throw new Error("Refusing to overwrite a non-empty cloud workspace with an empty snapshot.");
  }
  const nextVersion = Math.max(Number(current.version || 0) + 1, Date.now());
  const payload = {
    name: "Main Workspace",
    version: nextVersion,
    updated_at: new Date().toISOString(),
    updated_by: userId,
    state,
  };

  let data = null;
  let error = null;

  const updateAttempt = await supabase
    .from("workspaces")
    .update(payload)
    .eq("id", workspaceId)
    .select("id,version,updated_at")
    .maybeSingle();

  data = updateAttempt.data || null;
  error = updateAttempt.error || null;

  if (!data && !error) {
    const insertAttempt = await supabase
      .from("workspaces")
      .insert({ id: workspaceId, ...payload })
      .select("id,version,updated_at")
      .single();

    data = insertAttempt.data || null;
    error = insertAttempt.error || null;
  }

  if (error) throw error;
  if (!data) throw new Error("Cloud workspace save returned no row.");

  let backup = {
    available: false,
    saved: false,
    entry: null,
    notice: "",
  };

  try {
    backup = await createCloudWorkspaceBackup(state, {
      workspaceId,
      userId,
      version: Number(data.version || 0),
      reason: backupReason,
    });
  } catch (backupError) {
    backup = {
      available: true,
      saved: false,
      entry: null,
      notice: backupError instanceof Error ? backupError.message : "Cloud backup history failed.",
    };
  }

  return {
    id: data.id,
    version: Number(data.version || 0),
    updatedAt: data.updated_at || null,
    backup,
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
