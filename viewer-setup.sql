-- T1D Care Hub: view-only family access, e.g. grandparents (version 7)
-- Paste this whole file into Supabase: SQL Editor -> New query -> Run. Safe to run more than once.
--
-- Every family member is either a 'parent' (everything, as before) or a 'viewer' (view only).
-- Viewers can see: the care plan, sick-day status, and log entries from the last 36 hours.
-- Viewers cannot: add, change or delete anything, see older history, glucose files, reports, supplies,
-- appointments or backups, or see the family codes. Parents switch a view code on/off and can remove a viewer.

alter table public.family_members add column if not exists role text not null default 'parent';
do $$ begin
  alter table public.family_members drop constraint if exists family_members_role_chk;
  alter table public.family_members add constraint family_members_role_chk check (role in ('parent','viewer','carer'));
end $$;
alter table public.families add column if not exists viewer_code text unique;

-- "member" now means parent: every existing rule (records, backups, family codes) stays parents-only.
create or replace function public.is_member(fid uuid) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.family_members where family_id = fid and user_id = auth.uid() and role = 'parent');
$$;
create or replace function public.is_viewer(fid uuid) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.family_members where family_id = fid and user_id = auth.uid() and role = 'viewer');
$$;

drop policy if exists "read own membership" on public.family_members;
create policy "read own membership" on public.family_members for select to authenticated using (user_id = auth.uid());

-- read only: no insert, update or delete policies for viewers
drop policy if exists "viewers read" on public.records;
create policy "viewers read" on public.records for select to authenticated using (
  public.is_viewer(family_id) and (
       (collection = 'plan' and id = 'plan')
    or (collection = 'meta' and id = 'sickday')
    or (collection = 'log' and updated_at > now() - interval '36 hours')
  ));

-- parents join with the family code (max 4 parents)
create or replace function public.join_family(p_code text, p_name text) returns public.families
language plpgsql security definer set search_path = public as $$
declare f public.families;
begin
  if auth.uid() is null then raise exception 'Sign in first'; end if;
  if exists (select 1 from public.family_members where user_id = auth.uid()) then raise exception 'You are already in a family'; end if;
  select * into f from public.families where invite_code = upper(trim(p_code));
  if f.id is null then raise exception 'Family not found'; end if;
  if (select count(*) from public.family_members where family_id = f.id and role = 'parent') >= 4 then raise exception 'This family is full'; end if;
  insert into public.family_members (family_id, user_id, display_name, role) values (f.id, auth.uid(), left(p_name, 40), 'parent');
  return f;
end $$;

create or replace function public.set_viewer_code(p_on boolean) returns text
language plpgsql security definer set search_path = public as $$
declare fid uuid; c text;
begin
  select family_id into fid from public.family_members where user_id = auth.uid() and role = 'parent' limit 1;
  if fid is null then raise exception 'Only a parent can do this'; end if;
  c := case when p_on then 'V' || upper(substr(replace(gen_random_uuid()::text, '-', ''), 1, 7)) else null end;   -- built-in, works on Supabase
  update public.families set viewer_code = c where id = fid;
  return c;
end $$;

create or replace function public.join_as_viewer(p_code text, p_name text) returns uuid
language plpgsql security definer set search_path = public as $$
declare fid uuid;
begin
  if auth.uid() is null then raise exception 'Sign in first'; end if;
  if exists (select 1 from public.family_members where user_id = auth.uid()) then raise exception 'You are already in a family'; end if;
  select id into fid from public.families where viewer_code is not null and viewer_code = upper(trim(p_code));
  if fid is null then raise exception 'Code not found'; end if;
  if (select count(*) from public.family_members where family_id = fid and role = 'viewer') >= 6 then raise exception 'Too many viewers'; end if;
  insert into public.family_members (family_id, user_id, display_name, role) values (fid, auth.uid(), left(p_name, 40), 'viewer');
  return fid;
end $$;

create or replace function public.remove_viewer(p_user uuid) returns void
language plpgsql security definer set search_path = public as $$
declare fid uuid;
begin
  select family_id into fid from public.family_members where user_id = auth.uid() and role = 'parent' limit 1;
  if fid is null then raise exception 'Only a parent can do this'; end if;
  delete from public.family_members where family_id = fid and user_id = p_user and role = 'viewer';
end $$;

revoke all on function public.set_viewer_code(boolean)    from public, anon;
revoke all on function public.join_as_viewer(text, text)  from public, anon;
revoke all on function public.remove_viewer(uuid)         from public, anon;
grant execute on function public.set_viewer_code(boolean)   to authenticated;
grant execute on function public.join_as_viewer(text, text) to authenticated;
grant execute on function public.remove_viewer(uuid)        to authenticated;
grant execute on function public.join_family(text, text)    to authenticated;
