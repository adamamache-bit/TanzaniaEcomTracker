-- Migration 004: manual_ads_spend + extra_charges tables for Profit Center
-- Run in Supabase SQL Editor.

create table if not exists manual_ads_spend (
  id           uuid          primary key default gen_random_uuid(),
  week_start   date          not null,
  week_end     date          not null,
  product_id   text          not null default '',
  product_name text          not null default '',
  amount_usd   numeric(14,4) not null default 0,
  amount_tsh   numeric(14,2) not null default 0,
  notes        text          not null default '',
  created_at   timestamptz   not null default now()
);
alter table manual_ads_spend disable row level security;
create index if not exists idx_manual_ads_spend_week on manual_ads_spend (week_start);
create index if not exists idx_manual_ads_spend_product on manual_ads_spend (product_id) where product_id != '';

create table if not exists extra_charges (
  id          uuid          primary key default gen_random_uuid(),
  date        date          not null,
  category    text          not null default 'other',
  description text          not null default '',
  amount_usd  numeric(14,4) not null default 0,
  amount_tsh  numeric(14,2) not null default 0,
  created_at  timestamptz   not null default now()
);
alter table extra_charges disable row level security;
create index if not exists idx_extra_charges_date on extra_charges (date);
create index if not exists idx_extra_charges_category on extra_charges (category);
