create table if not exists public.workspaces (
  id text primary key,
  name text not null default 'Main Workspace',
  version bigint not null default 0,
  state jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default timezone('utc', now()),
  updated_by uuid references auth.users(id)
);

create table if not exists public.workspace_members (
  workspace_id text not null references public.workspaces(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null default 'admin',
  created_at timestamptz not null default timezone('utc', now()),
  primary key (workspace_id, user_id)
);

alter table public.workspaces enable row level security;
alter table public.workspace_members enable row level security;

create policy "members can read their workspace"
on public.workspaces
for select
to authenticated
using (
  exists (
    select 1
    from public.workspace_members wm
    where wm.workspace_id = workspaces.id
      and wm.user_id = auth.uid()
  )
);

create policy "members can update their workspace"
on public.workspaces
for update
to authenticated
using (
  exists (
    select 1
    from public.workspace_members wm
    where wm.workspace_id = workspaces.id
      and wm.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.workspace_members wm
    where wm.workspace_id = workspaces.id
      and wm.user_id = auth.uid()
  )
);

create policy "members can insert their workspace"
on public.workspaces
for insert
to authenticated
with check (
  exists (
    select 1
    from public.workspace_members wm
    where wm.workspace_id = workspaces.id
      and wm.user_id = auth.uid()
  )
);

create policy "users can read their memberships"
on public.workspace_members
for select
to authenticated
using (user_id = auth.uid());

create policy "users can insert their own membership"
on public.workspace_members
for insert
to authenticated
with check (user_id = auth.uid());

insert into public.workspaces (id, name, version, state)
values ('main-workspace', 'Main Workspace', 0, '{}'::jsonb)
on conflict (id) do nothing;

