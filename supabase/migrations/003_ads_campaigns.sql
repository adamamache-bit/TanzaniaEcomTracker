-- Migration 003: ads_campaigns table
-- Stores individual Meta campaign spend records for cumulative, per-product tracking.
-- Unique key: (workspace_id, campaign_id, date_start, date_end) — prevents double-counting
-- when the same date range is imported more than once.

create table if not exists ads_campaigns (
  id             bigint generated always as identity primary key,
  workspace_id   uuid        not null references workspaces(id) on delete cascade,
  campaign_id    text        not null,
  campaign_name  text        not null default '',
  date_start     date        not null,
  date_end       date        not null,
  spend_usd      numeric(14, 4) not null default 0,
  spend_tsh      numeric(14, 2) not null default 0,
  product_id     text        not null default '',
  is_mapped      boolean     not null default false,
  is_unmapped    boolean     not null default false,
  leads          integer     not null default 0,
  impressions    bigint      not null default 0,
  clicks         integer     not null default 0,
  imported_at    timestamptz not null default now(),
  updated_at     timestamptz not null default now(),

  constraint ads_campaigns_unique_key
    unique (workspace_id, campaign_id, date_start, date_end)
);

alter table ads_campaigns enable row level security;

create policy "workspace_members_ads_campaigns"
  on ads_campaigns for all
  to authenticated
  using (
    workspace_id in (
      select workspace_id from workspace_members where user_id = auth.uid()
    )
  )
  with check (
    workspace_id in (
      select workspace_id from workspace_members where user_id = auth.uid()
    )
  );

create index if not exists ads_campaigns_workspace_idx
  on ads_campaigns (workspace_id);

create index if not exists ads_campaigns_product_idx
  on ads_campaigns (workspace_id, product_id);

create index if not exists ads_campaigns_date_idx
  on ads_campaigns (workspace_id, date_start, date_end);
