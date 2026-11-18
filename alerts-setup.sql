-- T1D Care Hub: instant alerts when the app is closed (version 6)
-- Do the steps in README → "Instant alerts" first (deploy the parent-alert function and add its two secrets).
-- Then change the TWO lines marked EDIT below, paste this whole file into Supabase → SQL Editor → New query → Run.

create extension if not exists pg_net;

create or replace function public.t1d_alert_hook() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  fn_url text := 'https://YOUR-PROJECT.supabase.co/functions/v1/parent-alert';  -- EDIT: your Project URL + /functions/v1/parent-alert
  secret text := 'PASTE-YOUR-ALERT_SECRET-HERE';                                -- EDIT: the same text as the ALERT_SECRET secret
  was_on boolean := false;
begin
  if tg_op = 'UPDATE' then was_on := coalesce((old.data->>'active')::boolean, false); end if;
  if (tg_op = 'INSERT' and new.collection = 'log' and new.data->>'type' = 'hypo'
        and coalesce((new.data->>'at')::numeric, 0) > extract(epoch from now()) * 1000 - 1800000)   -- only fresh hypos, not restored old ones
     or (new.collection = 'meta' and new.id = 'sickday' and coalesce((new.data->>'active')::boolean, false) and not was_on) then
    perform net.http_post(
      url := fn_url,
      headers := jsonb_build_object('Content-Type', 'application/json', 'x-alert-secret', secret),
      body := jsonb_build_object('type', tg_op, 'table', 'records', 'record', to_jsonb(new),
                                 'old_record', case when tg_op = 'UPDATE' then to_jsonb(old) else null end));
  end if;
  return new;
end $$;

drop trigger if exists t1d_alert on public.records;
create trigger t1d_alert after insert or update on public.records
  for each row execute function public.t1d_alert_hook();
