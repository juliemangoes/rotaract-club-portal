-- ============================================================
-- Rotaract Club Portal — Supabase schema
-- Run this once in the SQL Editor of a fresh Supabase project.
-- ============================================================

create extension if not exists pgcrypto;

-- ---------- Tables ----------
create table public.clubs (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  join_code text unique not null default substr(md5(random()::text), 1, 8),
  created_at timestamptz not null default now()
);

create table public.club_members (
  club_id uuid not null references public.clubs(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null default 'member',
  created_at timestamptz not null default now(),
  primary key (club_id, user_id)
);

-- The whole club document (members, meetings, finances, projects…) lives here as JSONB,
-- versioned for optimistic concurrency. One row per club.
create table public.club_data (
  club_id uuid primary key references public.clubs(id) on delete cascade,
  doc jsonb not null default '{}'::jsonb,
  version bigint not null default 1,
  updated_at timestamptz not null default now()
);

create table public.push_subscriptions (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  club_id uuid not null references public.clubs(id) on delete cascade,
  subscription jsonb not null,
  created_at timestamptz not null default now(),
  unique (user_id, club_id)
);

-- ---------- Helper ----------
create or replace function public.is_club_member(c uuid)
returns boolean language sql stable security definer set search_path = public as
$$ select exists (select 1 from club_members where club_id = c and user_id = auth.uid()) $$;

-- ---------- Row-level security ----------
alter table public.clubs enable row level security;
alter table public.club_members enable row level security;
alter table public.club_data enable row level security;
alter table public.push_subscriptions enable row level security;

create policy "members read their club" on public.clubs
  for select using (public.is_club_member(id));

create policy "members read their memberships" on public.club_members
  for select using (user_id = auth.uid() or public.is_club_member(club_id));

create policy "members read club doc" on public.club_data
  for select using (public.is_club_member(club_id));

create policy "members manage own push subs" on public.push_subscriptions
  for all using (user_id = auth.uid()) with check (user_id = auth.uid() and public.is_club_member(club_id));

-- Writes to clubs / club_members / club_data go through the RPCs below (security definer),
-- so no direct insert/update policies are needed for them.

-- ---------- RPCs ----------
create or replace function public.create_club(p_name text)
returns uuid language plpgsql security definer set search_path = public as $$
declare cid uuid;
begin
  if auth.uid() is null then raise exception 'not signed in'; end if;
  insert into clubs (name) values (trim(p_name)) returning id into cid;
  insert into club_members (club_id, user_id, role) values (cid, auth.uid(), 'president');
  insert into club_data (club_id, doc) values (cid, '{}'::jsonb);
  return cid;
end $$;

create or replace function public.join_club(p_code text)
returns uuid language plpgsql security definer set search_path = public as $$
declare cid uuid;
begin
  if auth.uid() is null then raise exception 'not signed in'; end if;
  select id into cid from clubs where join_code = lower(trim(p_code));
  if cid is null then raise exception 'No club found for that invite code'; end if;
  insert into club_members (club_id, user_id, role) values (cid, auth.uid(), 'member')
    on conflict do nothing;
  return cid;
end $$;

-- Optimistic-concurrency save: returns the new version, or NULL if p_expected was stale.
create or replace function public.save_club_doc(p_club uuid, p_doc jsonb, p_expected bigint)
returns bigint language plpgsql security definer set search_path = public as $$
declare nv bigint;
begin
  if not public.is_club_member(p_club) then raise exception 'not a member of this club'; end if;
  update club_data set doc = p_doc, version = version + 1, updated_at = now()
    where club_id = p_club and version = p_expected
    returning version into nv;
  return nv; -- null when the version check failed (someone saved first)
end $$;

grant execute on function public.create_club(text) to authenticated;
grant execute on function public.join_club(text) to authenticated;
grant execute on function public.save_club_doc(uuid, jsonb, bigint) to authenticated;
grant execute on function public.is_club_member(uuid) to authenticated;

-- ---------- Realtime (live sync between members) ----------
alter publication supabase_realtime add table public.club_data;

-- ---------- File storage ----------
insert into storage.buckets (id, name, public) values ('club-files', 'club-files', true)
  on conflict (id) do nothing;

create policy "club members upload files" on storage.objects
  for insert to authenticated
  with check (bucket_id = 'club-files' and public.is_club_member((split_part(name, '/', 1))::uuid));

create policy "anyone reads club files" on storage.objects
  for select using (bucket_id = 'club-files');
