-- Migration 005: owner_injections table for Profit Center Cash Balance tab
-- Run in Supabase SQL Editor.

create table if not exists owner_injections (
  id          uuid          primary key default gen_random_uuid(),
  date        date          not null,
  amount_usd  numeric(14,4) not null default 0,
  amount_tsh  numeric(14,2) not null default 0,
  notes       text          not null default '',
  created_at  timestamptz   not null default now()
);
alter table owner_injections disable row level security;
create index if not exists idx_owner_injections_date on owner_injections (date);
