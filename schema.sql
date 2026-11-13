-- T1D Care Hub — database setup
-- Paste this whole file into Supabase: SQL Editor -> New query -> Run.
-- It is safe to run more than once.

create extension if not exists pgcrypto;

-- A family = the people who share one child's records
create table if not exists public.families (
  id          uuid primary key default gen_random_uuid(),
  invite_code text unique not null default upper(substr(encode(gen_random_bytes(6), 'hex'), 1, 8)),
  created_by  uuid references auth.users on delete set null,
  created_at  timestamptz not null default now()
);

create table if not exists public.family_members (
  family_id    uuid not null references public.families on delete cascade,
  user_id      uuid not null references auth.users on delete cascade,
  display_name text,
  created_at   timestamptz not null default now(),
  primary key (family_id, user_id)
);

-- Every saved item (log entry, CGM day, supply, measurement, care plan...) is one row
create table if not exists public.records (
  family_id  uuid not null references public.families on delete cascade,
  collection text not null,           -- log, cgm, supplies, growth, checks, meta, plan
  id         text not null,
  data       jsonb not null default '{}'::jsonb,
  updated_by uuid,
  updated_at timestamptz not null default now(),
  primary key (family_id, collection, id)
);
create index if not exists records_family_idx on public.records (family_id, collection);

-- Is the signed-in user in this family?
create or replace function public.is_member(fid uuid) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.family_members where family_id = fid and user_id = auth.uid());
$$;

alter table public.families       enable row level security;
alter table public.family_members enable row level security;
alter table public.records        enable row level security;

drop policy if exists "members read family"   on public.families;
drop policy if exists "members read members"  on public.family_members;
drop policy if exists "members use records"   on public.records;
create policy "members read family"  on public.families       for select to authenticated using (public.is_member(id));
create policy "members read members" on public.family_members for select to authenticated using (public.is_member(family_id));
create policy "members use records"  on public.records        for all    to authenticated using (public.is_member(family_id)) with check (public.is_member(family_id));

-- First parent: create a family and join it
create or replace function public.create_family(p_name text) returns public.families
language plpgsql security definer set search_path = public as $$
declare f public.families;
begin
  if auth.uid() is null then raise exception 'Sign in first'; end if;
  if exists (select 1 from public.family_members where user_id = auth.uid()) then
    raise exception 'You are already in a family';
  end if;
  insert into public.families (created_by) values (auth.uid()) returning * into f;
  insert into public.family_members (family_id, user_id, display_name) values (f.id, auth.uid(), left(p_name, 40));
  return f;
end $$;

-- Second parent: join with the family code (max 4 people per family)
create or replace function public.join_family(p_code text, p_name text) returns public.families
language plpgsql security definer set search_path = public as $$
declare f public.families;
begin
  if auth.uid() is null then raise exception 'Sign in first'; end if;
  select * into f from public.families where invite_code = upper(trim(p_code));
  if f.id is null then raise exception 'Family not found'; end if;
  if (select count(*) from public.family_members where family_id = f.id) >= 4 then
    raise exception 'This family is full';
  end if;
  insert into public.family_members (family_id, user_id, display_name)
  values (f.id, auth.uid(), left(p_name, 40)) on conflict do nothing;
  return f;
end $$;

revoke all on function public.create_family(text)     from public, anon;
revoke all on function public.join_family(text, text) from public, anon;
grant execute on function public.create_family(text)     to authenticated;
grant execute on function public.join_family(text, text) to authenticated;

-- Live sync between phones
alter table public.records replica identity full;
do $$ begin
  alter publication supabase_realtime add table public.records;
exception when duplicate_object then null; end $$;
