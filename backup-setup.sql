-- T1D Care Hub — automatic daily backups
-- Paste this whole file into Supabase: SQL Editor -> New query -> Run.
-- It is safe to run more than once. Run schema.sql first (you already have).
--
-- What it does:
--   * every night at 02:30 (UTC) Supabase copies all of your family's records into a "backups" table
--   * the app also makes one the first time either of you opens it each day, if the night one is missing
--   * the last 30 days are kept; older ones are deleted automatically
--   * either parent can restore any of them from the Backup tab

create table if not exists public.backups (
  family_id  uuid not null references public.families on delete cascade,
  day        date not null,
  created_at timestamptz not null default now(),
  items      int  not null default 0,
  data       jsonb not null,
  primary key (family_id, day)
);
alter table public.backups enable row level security;
drop policy if exists "members read backups" on public.backups;
create policy "members read backups" on public.backups for select to authenticated using (public.is_member(family_id));

-- copy one family's records into today's backup (UK date), keep 30 days
create or replace function public.snapshot_family(fid uuid) returns int
language plpgsql security definer set search_path = public as $$
declare n int; d date := (now() at time zone 'Europe/London')::date;
begin
  insert into public.backups (family_id, day, items, data)
  select fid, d, count(*), coalesce(jsonb_agg(jsonb_build_object('c', collection, 'i', id, 'd', data)), '[]'::jsonb)
    from public.records where family_id = fid and not (collection = 'meta' and id = 'backup')
  on conflict (family_id, day) do update set items = excluded.items, data = excluded.data, created_at = now()
  returning items into n;
  -- tell the app when the last backup was made
  insert into public.records (family_id, collection, id, data, updated_at)
  values (fid, 'meta', 'backup', jsonb_build_object('at', (extract(epoch from now()) * 1000)::bigint, 'auto', true, 'items', n), now())
  on conflict (family_id, collection, id) do update set data = excluded.data, updated_at = now();
  delete from public.backups where family_id = fid and day < d - 30;
  return n;
end $$;

-- the app calls this ("Back up now", and once a day when it opens)
create or replace function public.backup_now() returns int
language plpgsql security definer set search_path = public as $$
declare fid uuid;
begin
  select family_id into fid from public.family_members where user_id = auth.uid() limit 1;
  if fid is null then raise exception 'Not in a family'; end if;
  return public.snapshot_family(fid);
end $$;

-- the nightly job calls this
create or replace function public.backup_all_families() returns int
language plpgsql security definer set search_path = public as $$
declare f record; k int := 0;
begin
  for f in select id from public.families loop perform public.snapshot_family(f.id); k := k + 1; end loop;
  return k;
end $$;

revoke execute on function public.snapshot_family(uuid)     from public, anon, authenticated;
revoke execute on function public.backup_all_families()     from public, anon, authenticated;
revoke execute on function public.backup_now()              from public, anon;
grant  execute on function public.backup_now()              to authenticated;

-- nightly schedule (pg_cron is included free in every Supabase project)
create extension if not exists pg_cron;
select cron.schedule('t1d-daily-backup', '30 2 * * *', 'select public.backup_all_families()');
