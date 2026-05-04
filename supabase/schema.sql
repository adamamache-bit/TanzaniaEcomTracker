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

create table if not exists public.workspace_backups (
  id bigint generated always as identity primary key,
  workspace_id text not null references public.workspaces(id) on delete cascade,
  workspace_version bigint not null default 0,
  state jsonb not null default '{}'::jsonb,
  summary jsonb not null default '{}'::jsonb,
  reason text not null default 'autosave',
  created_at timestamptz not null default timezone('utc', now()),
  created_by uuid references auth.users(id)
);

create index if not exists workspace_backups_workspace_created_at_idx
on public.workspace_backups (workspace_id, created_at desc);

create index if not exists workspace_backups_workspace_version_idx
on public.workspace_backups (workspace_id, workspace_version desc);

alter table public.workspaces enable row level security;
alter table public.workspace_members enable row level security;
alter table public.workspace_backups enable row level security;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspaces'
      and policyname = 'members can read their workspace'
  ) then
    execute $policy$
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
      )
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspaces'
      and policyname = 'members can update their workspace'
  ) then
    execute $policy$
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
      )
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspaces'
      and policyname = 'members can insert their workspace'
  ) then
    execute $policy$
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
      )
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspace_members'
      and policyname = 'users can read their memberships'
  ) then
    execute $policy$
      create policy "users can read their memberships"
      on public.workspace_members
      for select
      to authenticated
      using (user_id = auth.uid())
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspace_members'
      and policyname = 'users can insert their own membership'
  ) then
    execute $policy$
      create policy "users can insert their own membership"
      on public.workspace_members
      for insert
      to authenticated
      with check (user_id = auth.uid())
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspace_backups'
      and policyname = 'members can read workspace backups'
  ) then
    execute $policy$
      create policy "members can read workspace backups"
      on public.workspace_backups
      for select
      to authenticated
      using (
        exists (
          select 1
          from public.workspace_members wm
          where wm.workspace_id = workspace_backups.workspace_id
            and wm.user_id = auth.uid()
        )
      )
    $policy$;
  end if;
end
$$;

do $$
begin
  if not exists (
    select 1
    from pg_policies
    where schemaname = 'public'
      and tablename = 'workspace_backups'
      and policyname = 'members can insert workspace backups'
  ) then
    execute $policy$
      create policy "members can insert workspace backups"
      on public.workspace_backups
      for insert
      to authenticated
      with check (
        exists (
          select 1
          from public.workspace_members wm
          where wm.workspace_id = workspace_backups.workspace_id
            and wm.user_id = auth.uid()
        )
      )
    $policy$;
  end if;
end
$$;

insert into public.workspaces (id, name, version, state)
values ('main-workspace', 'Main Workspace', 0, '{}'::jsonb)
on conflict (id) do nothing;
