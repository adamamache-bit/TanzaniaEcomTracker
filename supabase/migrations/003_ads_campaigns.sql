-- Migration 003: ads_campaigns table
-- Simple table for storing Meta ads campaign spend records.
-- No workspace_id — single-tenant app, RLS disabled for simplicity.
-- Unique key: (campaign_id, date_start) — prevents double-counting same campaign/day.
--
-- Run this in the Supabase SQL editor:
--   1. Open your Supabase project dashboard
--   2. Go to SQL Editor
--   3. Paste and run this entire file

create table if not exists ads_campaigns (
  id             uuid          primary key default gen_random_uuid(),
  campaign_id    text          not null default '',
  campaign_name  text          not null default '',
  spend_usd      numeric(14, 4) not null default 0,
  spend_tsh      numeric(14, 2) not null default 0,
  date_start     date          not null,
  date_end       date          not null,
  product_id     text          not null default '',
  product_name   text          not null default '',
  is_mapped      boolean       not null default false,
  impressions    numeric       not null default 0,
  clicks         numeric       not null default 0,
  leads          numeric       not null default 0,
  created_at     timestamptz   not null default now(),
  updated_at     timestamptz   not null default now()
);

-- Disable RLS so the anon key can read/write without auth
alter table ads_campaigns disable row level security;

-- Unique constraint prevents duplicate import of same campaign+period
create unique index if not exists ads_campaigns_campaign_date_idx
  on ads_campaigns (campaign_id, date_start)
  where campaign_id != '';

create index if not exists ads_campaigns_product_idx
  on ads_campaigns (product_id)
  where product_id != '';
