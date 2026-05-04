-- Tanzania E-Commerce Tracker — Supabase schema
-- Run this once in the Supabase SQL Editor for your project.
-- Project: https://gyzioitptnnhumpnzift.supabase.co

-- ============================================================
-- TABLE: workspaces
-- Stores the entire app state as a versioned JSON blob.
-- One row per workspace (workspace_id = "main-workspace").
-- ============================================================
CREATE TABLE IF NOT EXISTS workspaces (
  id           TEXT PRIMARY KEY,
  name         TEXT,
  version      BIGINT NOT NULL DEFAULT 0,
  updated_at   TIMESTAMPTZ DEFAULT NOW(),
  updated_by   UUID,
  state        JSONB
);

-- ============================================================
-- TABLE: workspace_members
-- Tracks which Supabase auth users belong to which workspace.
-- ============================================================
CREATE TABLE IF NOT EXISTS workspace_members (
  workspace_id TEXT NOT NULL,
  user_id      UUID NOT NULL,
  role         TEXT NOT NULL DEFAULT 'admin',
  joined_at    TIMESTAMPTZ DEFAULT NOW(),
  PRIMARY KEY (workspace_id, user_id)
);

-- ============================================================
-- TABLE: workspace_backups
-- Auto-versioned snapshots created on every save.
-- Used for cloud restore history.
-- ============================================================
CREATE TABLE IF NOT EXISTS workspace_backups (
  id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  workspace_id       TEXT NOT NULL,
  workspace_version  BIGINT DEFAULT 0,
  state              JSONB,
  created_at         TIMESTAMPTZ DEFAULT NOW(),
  created_by         UUID,
  reason             TEXT,
  summary            JSONB
);

-- ============================================================
-- RLS: Enable row-level security with permissive policies.
-- All authenticated users can read/write (private app, 2 users).
-- ============================================================
ALTER TABLE workspaces        ENABLE ROW LEVEL SECURITY;
ALTER TABLE workspace_members ENABLE ROW LEVEL SECURITY;
ALTER TABLE workspace_backups ENABLE ROW LEVEL SECURITY;

-- Drop existing policies if re-running this script
DROP POLICY IF EXISTS "authenticated_all" ON workspaces;
DROP POLICY IF EXISTS "authenticated_all" ON workspace_members;
DROP POLICY IF EXISTS "authenticated_all" ON workspace_backups;

CREATE POLICY "authenticated_all" ON workspaces
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "authenticated_all" ON workspace_members
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "authenticated_all" ON workspace_backups
  FOR ALL TO authenticated USING (true) WITH CHECK (true);

-- ============================================================
-- INDEX: Speed up backup listing queries
-- ============================================================
CREATE INDEX IF NOT EXISTS idx_workspace_backups_workspace_created
  ON workspace_backups (workspace_id, created_at DESC);

-- ============================================================
-- Real-time: Enable replication on workspaces table
-- (required for subscribeToCloudWorkspace to work)
-- ============================================================
ALTER PUBLICATION supabase_realtime ADD TABLE workspaces;
