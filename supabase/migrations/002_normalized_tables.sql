-- Tanzania E-Commerce Tracker — Normalized tables
-- Run this in the Supabase SQL Editor after 001_workspace_tables.sql
-- These tables mirror the workspace JSON blob for direct querying.

-- ============================================================
-- TABLE: products
-- ============================================================
CREATE TABLE IF NOT EXISTS products (
  id                      TEXT NOT NULL,
  workspace_id            TEXT NOT NULL DEFAULT 'main-workspace',
  name                    TEXT DEFAULT '',
  source                  TEXT DEFAULT '',
  selling_price           NUMERIC DEFAULT 0,
  purchase_unit_price     NUMERIC DEFAULT 0,
  total_qty               INTEGER DEFAULT 0,
  shipping_total          NUMERIC DEFAULT 0,
  other_charges           NUMERIC DEFAULT 0,
  delivery                NUMERIC DEFAULT 0,
  estimated_arrival_days  INTEGER DEFAULT 0,
  stock_arrival_status    TEXT DEFAULT '',
  stock_ordered_at        DATE,
  next_arrival_check_date DATE,
  stock_arrived_at        DATE,
  offers                  JSONB DEFAULT '[]',
  extra                   JSONB DEFAULT '{}',
  updated_at              TIMESTAMPTZ DEFAULT NOW(),
  PRIMARY KEY (id, workspace_id)
);

-- ============================================================
-- TABLE: orders  (customers / leads / shipping combined)
-- ============================================================
CREATE TABLE IF NOT EXISTS orders (
  id                         TEXT NOT NULL,
  workspace_id               TEXT NOT NULL DEFAULT 'main-workspace',
  customer_name              TEXT DEFAULT '',
  phone                      TEXT DEFAULT '',
  city                       TEXT DEFAULT '',
  address                    TEXT DEFAULT '',
  product_id                 TEXT DEFAULT '',
  quantity                   INTEGER DEFAULT 1,
  order_date                 DATE,
  payment_method             TEXT DEFAULT 'COD',
  notes                      TEXT DEFAULT '',
  lead_source                TEXT DEFAULT 'facebook',
  campaign_name              TEXT DEFAULT '',
  adset_name                 TEXT DEFAULT '',
  creative_name              TEXT DEFAULT '',
  priority                   TEXT DEFAULT 'normal',
  customer_type              TEXT DEFAULT 'new',
  call_attempts              INTEGER DEFAULT 0,
  cancel_reason              TEXT DEFAULT '',
  unreached_reason           TEXT DEFAULT '',
  carrier_name               TEXT DEFAULT '',
  tracking_number            TEXT DEFAULT '',
  expected_delivery_date     DATE,
  actual_delivery_date       DATE,
  return_reason              TEXT DEFAULT '',
  order_total_tzs            NUMERIC DEFAULT 0,
  amount_tsh                 NUMERIC DEFAULT 0,
  amount_usd                 NUMERIC DEFAULT 0,
  source_order_id            TEXT,
  import_source              TEXT,
  last_imported_at           TIMESTAMPTZ,
  last_shipping_imported_at  TIMESTAMPTZ,
  updated_at                 TIMESTAMPTZ,
  confirmation_updated_at    TIMESTAMPTZ,
  assigned_to                TEXT DEFAULT '',
  confirmation_status        TEXT DEFAULT '',
  shipping_status            TEXT DEFAULT '',
  status                     TEXT DEFAULT '',
  history                    JSONB DEFAULT '[]',
  extra                      JSONB DEFAULT '{}',
  PRIMARY KEY (id, workspace_id)
);

-- ============================================================
-- TABLE: tracking  (ad spend rows, Meta import rows)
-- ============================================================
CREATE TABLE IF NOT EXISTS tracking (
  id               TEXT NOT NULL,
  workspace_id     TEXT NOT NULL DEFAULT 'main-workspace',
  product_id       TEXT DEFAULT '',
  ad_spend         NUMERIC DEFAULT 0,
  orders           INTEGER DEFAULT 0,
  confirmed        INTEGER DEFAULT 0,
  delivered        INTEGER DEFAULT 0,
  name             TEXT DEFAULT '',
  date_start       DATE,
  date_end         DATE,
  meta_managed     BOOLEAN DEFAULT false,
  meta_imported_at TIMESTAMPTZ,
  meta_currency    TEXT DEFAULT 'USD',
  extra            JSONB DEFAULT '{}',
  updated_at       TIMESTAMPTZ DEFAULT NOW(),
  PRIMARY KEY (id, workspace_id)
);

-- ============================================================
-- TABLE: settings  (key-value: serviceForm, situationData, etc.)
-- ============================================================
CREATE TABLE IF NOT EXISTS settings (
  key          TEXT NOT NULL,
  workspace_id TEXT NOT NULL DEFAULT 'main-workspace',
  value        JSONB,
  updated_at   TIMESTAMPTZ DEFAULT NOW(),
  PRIMARY KEY (key, workspace_id)
);

-- ============================================================
-- TABLE: audit_trail  (order history entries)
-- ============================================================
CREATE TABLE IF NOT EXISTS audit_trail (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  workspace_id  TEXT NOT NULL DEFAULT 'main-workspace',
  customer_id   TEXT DEFAULT '',
  customer_name TEXT DEFAULT '',
  product_id    TEXT DEFAULT '',
  action        TEXT DEFAULT '',
  source        TEXT DEFAULT 'system',
  details       TEXT DEFAULT '',
  occurred_at   TIMESTAMPTZ,
  created_at    TIMESTAMPTZ DEFAULT NOW()
);

-- ============================================================
-- RLS: Disabled (private app, 2 users)
-- ============================================================
ALTER TABLE products    DISABLE ROW LEVEL SECURITY;
ALTER TABLE orders      DISABLE ROW LEVEL SECURITY;
ALTER TABLE tracking    DISABLE ROW LEVEL SECURITY;
ALTER TABLE settings    DISABLE ROW LEVEL SECURITY;
ALTER TABLE audit_trail DISABLE ROW LEVEL SECURITY;

-- ============================================================
-- Indexes
-- ============================================================
CREATE INDEX IF NOT EXISTS idx_orders_workspace      ON orders (workspace_id);
CREATE INDEX IF NOT EXISTS idx_orders_product        ON orders (workspace_id, product_id);
CREATE INDEX IF NOT EXISTS idx_orders_status         ON orders (workspace_id, status);
CREATE INDEX IF NOT EXISTS idx_products_workspace    ON products (workspace_id);
CREATE INDEX IF NOT EXISTS idx_tracking_workspace    ON tracking (workspace_id);
CREATE INDEX IF NOT EXISTS idx_tracking_product      ON tracking (workspace_id, product_id);
CREATE INDEX IF NOT EXISTS idx_audit_workspace_time  ON audit_trail (workspace_id, occurred_at DESC);
