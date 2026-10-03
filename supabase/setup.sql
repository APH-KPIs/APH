-- =====================================================================
-- لوحة مؤشرات الأداء — مستشفى إرادة والصحة النفسية بأبها
-- Supabase setup: run once in the Supabase SQL Editor (safe to re-run).
--
-- Model
--   * Reads are public (same as the dashboard's guest mode): kpis, records,
--     target_years and app_events can be selected with the anon key.
--   * Every write goes through a SECURITY DEFINER function that checks the
--     caller's session token and role. Tables have no write policies.
--   * Passwords are bcrypt hashes (pgcrypto); sessions expire after 12 hours.
--   * Each write inserts a row in app_events; open dashboards subscribe to it
--     (Realtime) and pull the changes.
-- =====================================================================

create schema if not exists extensions;
create extension if not exists pgcrypto with schema extensions;

-- ---------- tables ----------
create table if not exists public.app_users (
    username        text primary key check (username ~ '^[a-z0-9._-]{3,32}$'),
    display_name    text not null,
    role            text not null check (role in ('admin','manager','data')),
    active          boolean not null default true,
    pass_hash       text not null,
    must_change     boolean not null default true,
    created_at      timestamptz not null default now(),
    updated_at      timestamptz not null default now()
);

create table if not exists public.app_sessions (
    token       uuid primary key default gen_random_uuid(),
    username    text not null references public.app_users(username) on delete cascade,
    expires_at  timestamptz not null default now() + interval '12 hours'
);

create table if not exists public.kpis (
    code        text primary key,
    data        jsonb not null,
    updated_at  timestamptz not null default clock_timestamp()
);

-- One reading per KPI and date: id = kpi_code || '|' || date.
create table if not exists public.records (
    id           text primary key,
    kpi_code     text not null,
    date         date not null,
    year         int not null,
    quarter      text not null check (quarter in ('Q1','Q2','Q3','Q4')),
    actual       double precision,
    numerator    double precision,
    denominator  double precision,
    weight       double precision,
    data_source  text,
    notes        text,
    dq           text,
    approved     boolean not null default false,
    created_by   text,
    created_user text,
    created_at   timestamptz not null default now(),
    updated_by   text,
    updated_at   timestamptz not null default clock_timestamp(),
    approved_by  text,
    approved_at  timestamptz,
    deleted      boolean not null default false
);
create index if not exists records_updated_idx on public.records (updated_at);
-- Change cursor for the dashboards: every insert/update/delete takes a new rev, and pages pull rev > last seen.
create sequence if not exists public.records_rev_seq;
alter table public.records add column if not exists rev bigint not null default nextval('public.records_rev_seq');
create index if not exists records_rev_idx on public.records (rev);
create index if not exists records_kpi_year_idx on public.records (kpi_code, year);

create table if not exists public.target_years (
    year        int primary key,
    data        jsonb not null,
    updated_at  timestamptz not null default clock_timestamp()
);

create table if not exists public.audit_log (
    id            bigserial primary key,
    at            timestamptz not null default now(),
    username      text,
    display_name  text,
    action        text not null,
    item          text,
    old_value     text,
    new_value     text,
    reason        text
);

create table if not exists public.app_events (
    id      bigserial primary key,
    kind    text not null,
    at      timestamptz not null default now(),
    by_user text
);

-- ---------- row level security: public read, no direct writes ----------
alter table public.app_users    enable row level security;
alter table public.app_sessions enable row level security;
alter table public.kpis         enable row level security;
alter table public.records      enable row level security;
alter table public.target_years enable row level security;
alter table public.audit_log    enable row level security;
alter table public.app_events   enable row level security;

drop policy if exists kpis_read on public.kpis;
create policy kpis_read on public.kpis for select using (true);
drop policy if exists records_read on public.records;
create policy records_read on public.records for select using (true);
drop policy if exists target_years_read on public.target_years;
create policy target_years_read on public.target_years for select using (true);
drop policy if exists app_events_read on public.app_events;
create policy app_events_read on public.app_events for select using (true);

-- ---------- helpers ----------
create or replace function public.app_session_user(p_token uuid)
returns public.app_users language plpgsql security definer set search_path = public as $$
declare u public.app_users;
begin
    select a.* into u from public.app_sessions s join public.app_users a on a.username = s.username
     where s.token = p_token and s.expires_at > now() and a.active;
    if not found then raise exception 'SESSION_EXPIRED' using errcode = '28000'; end if;
    return u;
end $$;

create or replace function public.app_require(p_token uuid, p_roles text[])
returns public.app_users language plpgsql security definer set search_path = public as $$
declare u public.app_users;
begin
    u := public.app_session_user(p_token);
    if not (u.role = any(p_roles)) then raise exception 'FORBIDDEN' using errcode = '42501'; end if;
    return u;
end $$;

create or replace function public.app_bump(p_kind text, p_user text)
returns void language sql security definer set search_path = public as $$
    insert into public.app_events(kind, by_user) values (p_kind, p_user);
    delete from public.app_events where at < now() - interval '2 days';
$$;

create or replace function public.app_user_json(u public.app_users)
returns jsonb language sql immutable as $$
    select jsonb_build_object('username', u.username, 'displayName', u.display_name, 'role', u.role,
                              'active', u.active, 'mustChange', u.must_change, 'createdAt', u.created_at);
$$;

-- ---------- authentication ----------
-- Names shown on the login screen's account picker (no secrets).
create or replace function public.public_users()
returns jsonb language sql security definer set search_path = public as $$
    select coalesce(jsonb_agg(jsonb_build_object('username', username, 'displayName', display_name, 'role', role)
                    order by case role when 'admin' then 0 when 'manager' then 1 else 2 end, display_name), '[]'::jsonb)
      from public.app_users where active;
$$;

create or replace function public.login(p_username text, p_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare u public.app_users; t uuid;
begin
    select * into u from public.app_users where username = lower(trim(p_username)) and active;
    if not found or u.pass_hash <> extensions.crypt(p_password, u.pass_hash) then
        perform pg_sleep(0.5);  -- slow down guessing
        raise exception 'BAD_CREDENTIALS' using errcode = '28P01';
    end if;
    delete from public.app_sessions where expires_at < now();
    insert into public.app_sessions(username) values (u.username) returning token into t;
    return jsonb_build_object('token', t, 'user', public.app_user_json(u));
end $$;

create or replace function public.logout(p_token uuid)
returns void language sql security definer set search_path = public as $$
    delete from public.app_sessions where token = p_token;
$$;

create or replace function public.whoami(p_token uuid)
returns jsonb language plpgsql security definer set search_path = public as $$
begin
    return public.app_user_json(public.app_session_user(p_token));
end $$;

create or replace function public.change_password(p_token uuid, p_old text, p_new text)
returns void language plpgsql security definer set search_path = public, extensions as $$
declare u public.app_users;
begin
    u := public.app_session_user(p_token);
    if u.pass_hash <> extensions.crypt(p_old, u.pass_hash) then raise exception 'BAD_CREDENTIALS' using errcode = '28P01'; end if;
    if length(coalesce(p_new, '')) < 8 then raise exception 'WEAK_PASSWORD' using errcode = '22023'; end if;
    update public.app_users set pass_hash = extensions.crypt(p_new, extensions.gen_salt('bf')), must_change = false, updated_at = now()
     where username = u.username;
    insert into public.audit_log(username, display_name, action, item, reason) values (u.username, u.display_name, 'CHANGE_PASSWORD', u.username, 'تغيير كلمة المرور');
end $$;

-- ---------- users (admin only) ----------
create or replace function public.list_users(p_token uuid)
returns jsonb language plpgsql security definer set search_path = public as $$
begin
    perform public.app_require(p_token, array['admin']);
    return coalesce((select jsonb_agg(public.app_user_json(a) order by a.created_at) from public.app_users a), '[]'::jsonb);
end $$;

-- Creates the user, or updates name/role/active; p_password resets the password (user must change it).
create or replace function public.save_user(p_token uuid, p_username text, p_display_name text, p_role text, p_active boolean, p_password text)
returns jsonb language plpgsql security definer set search_path = public, extensions as $$
declare me public.app_users; uname text := lower(trim(p_username)); existing public.app_users; result public.app_users;
begin
    me := public.app_require(p_token, array['admin']);
    if coalesce(trim(p_display_name), '') = '' then raise exception 'NAME_REQUIRED' using errcode = '22023'; end if;
    if p_role not in ('admin','manager','data') then raise exception 'BAD_ROLE' using errcode = '22023'; end if;
    if p_password is not null and length(p_password) < 8 then raise exception 'WEAK_PASSWORD' using errcode = '22023'; end if;
    select * into existing from public.app_users where username = uname;
    if found then
        -- Never leave the system without an active admin.
        if existing.role = 'admin' and (p_role <> 'admin' or not p_active)
           and (select count(*) from public.app_users where role = 'admin' and active and username <> uname) = 0 then
            raise exception 'LAST_ADMIN' using errcode = '22023';
        end if;
        update public.app_users set display_name = trim(p_display_name), role = p_role, active = p_active,
               pass_hash = case when p_password is null then pass_hash else extensions.crypt(p_password, extensions.gen_salt('bf')) end,
               must_change = case when p_password is null then must_change else true end, updated_at = now()
         where username = uname returning * into result;
        if not p_active or p_password is not null then delete from public.app_sessions where username = uname; end if;
        insert into public.audit_log(username, display_name, action, item, old_value, new_value, reason)
        values (me.username, me.display_name, 'EDIT_USER', uname, existing.role || case when existing.active then '' else ' (معطل)' end,
                p_role || case when p_active then '' else ' (معطل)' end, case when p_password is null then 'تعديل مستخدم' else 'تعديل مستخدم وإعادة تعيين كلمة المرور' end);
    else
        if p_password is null then raise exception 'PASSWORD_REQUIRED' using errcode = '22023'; end if;
        insert into public.app_users(username, display_name, role, active, pass_hash, must_change)
        values (uname, trim(p_display_name), p_role, p_active, extensions.crypt(p_password, extensions.gen_salt('bf')), true)
        returning * into result;
        insert into public.audit_log(username, display_name, action, item, new_value, reason)
        values (me.username, me.display_name, 'ADD_USER', uname, p_role, 'إضافة مستخدم');
    end if;
    perform public.app_bump('users', me.username);
    return public.app_user_json(result);
end $$;

-- ---------- readings ----------
-- Rules: data entry saves pending readings and may only change its own pending ones;
-- managers/admins may save, approve and overwrite any reading.
create or replace function public.upsert_records(p_token uuid, p_rows jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; r jsonb; ex public.records; approver boolean; saved int := 0; skipped int := 0; want_approved boolean; rid text;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    approver := me.role in ('admin','manager');
    for r in select * from jsonb_array_elements(coalesce(p_rows, '[]'::jsonb)) loop
        rid := (r->>'kpiCode') || '|' || (r->>'date');
        select * into ex from public.records where id = rid;
        if found and not ex.deleted and not approver and (ex.approved or ex.created_user is distinct from me.username) then
            skipped := skipped + 1; continue;
        end if;
        want_approved := approver and coalesce((r->>'approved')::boolean, false);
        insert into public.records as t (id, kpi_code, date, year, quarter, actual, numerator, denominator, weight, data_source, notes, dq,
                                         approved, created_by, created_user, created_at, updated_by, updated_at, approved_by, approved_at, deleted)
        values (rid, r->>'kpiCode', (r->>'date')::date, (r->>'year')::int, r->>'quarter',
                (r->>'actualValue')::double precision, (r->>'numerator')::double precision, (r->>'denominator')::double precision,
                (r->>'weight')::double precision, r->>'dataSource', r->>'notes', nullif(r->>'dq', ''),
                want_approved, coalesce(r->>'createdBy', me.display_name), me.username, now(), me.display_name, clock_timestamp(),
                case when want_approved then me.display_name end, case when want_approved then now() end, false)
        on conflict (id) do update set
            year = excluded.year, quarter = excluded.quarter, actual = excluded.actual, numerator = excluded.numerator,
            denominator = excluded.denominator, weight = excluded.weight, data_source = excluded.data_source,
            notes = excluded.notes, dq = excluded.dq, approved = excluded.approved,
            created_by = case when t.deleted then excluded.created_by else t.created_by end,
            created_user = case when t.deleted then excluded.created_user else t.created_user end,
            updated_by = me.display_name, updated_at = clock_timestamp(), deleted = false, rev = nextval('public.records_rev_seq'),
            approved_by = case when excluded.approved and not t.approved then me.display_name when excluded.approved then t.approved_by end,
            approved_at = case when excluded.approved and not t.approved then now() when excluded.approved then t.approved_at end;
        saved := saved + 1;
    end loop;
    if saved > 0 then perform public.app_bump('records', me.username); end if;
    return jsonb_build_object('saved', saved, 'skipped', skipped);
end $$;

create or replace function public.delete_records(p_token uuid, p_ids text[])
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; n int;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    update public.records set deleted = true, updated_at = clock_timestamp(), updated_by = me.display_name, rev = nextval('public.records_rev_seq')
     where id = any(p_ids) and not deleted
       and (me.role in ('admin','manager') or (not approved and created_user = me.username));
    get diagnostics n = row_count;
    if n > 0 then perform public.app_bump('records', me.username); end if;
    return jsonb_build_object('deleted', n);
end $$;

-- Bulk import (Excel). p_replace_codes: approvers may first remove existing readings of these KPIs.
create or replace function public.import_records(p_token uuid, p_rows jsonb, p_replace_codes text[])
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; removed int := 0; res jsonb;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    if p_replace_codes is not null and array_length(p_replace_codes, 1) > 0 and me.role in ('admin','manager') then
        update public.records set deleted = true, updated_at = clock_timestamp(), updated_by = me.display_name, rev = nextval('public.records_rev_seq')
         where kpi_code = any(p_replace_codes) and not deleted;
        get diagnostics removed = row_count;
    end if;
    res := public.upsert_records(p_token, p_rows);
    return res || jsonb_build_object('removed', removed);
end $$;

-- ---------- KPI definitions (admin only) ----------
-- p_replace: the list becomes the official list; KPIs not in it (and their readings) are removed.
create or replace function public.save_kpis(p_token uuid, p_kpis jsonb, p_replace boolean)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; k jsonb; codes text[]; n int := 0;
begin
    me := public.app_require(p_token, array['admin']);
    select coalesce(array_agg(x->>'code'), '{}') into codes from jsonb_array_elements(p_kpis) x;
    for k in select * from jsonb_array_elements(p_kpis) loop
        insert into public.kpis(code, data, updated_at) values (k->>'code', k, clock_timestamp())
        on conflict (code) do update set data = excluded.data, updated_at = clock_timestamp();
        n := n + 1;
    end loop;
    if p_replace then
        delete from public.kpis where not (code = any(codes));
        update public.records set deleted = true, updated_at = clock_timestamp(), updated_by = me.display_name, rev = nextval('public.records_rev_seq')
         where not deleted and not (kpi_code = any(codes));
    end if;
    perform public.app_bump('kpis', me.username);
    return jsonb_build_object('saved', n);
end $$;

create or replace function public.delete_kpi(p_token uuid, p_code text)
returns void language plpgsql security definer set search_path = public as $$
declare me public.app_users;
begin
    me := public.app_require(p_token, array['admin']);
    delete from public.kpis where code = p_code;
    perform public.app_bump('kpis', me.username);
end $$;

-- Targets and directions (admin and performance manager). Only those two fields of existing KPIs change;
-- a KPI not yet in the table (fresh database) is stored as sent.
create or replace function public.save_kpi_targets(p_token uuid, p_kpis jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; k jsonb; n int := 0;
begin
    me := public.app_require(p_token, array['admin','manager']);
    for k in select * from jsonb_array_elements(p_kpis) loop
        insert into public.kpis(code, data, updated_at) values (k->>'code', k, clock_timestamp())
        on conflict (code) do update set
            data = public.kpis.data || jsonb_build_object('targets', coalesce(k->'targets', '{}'::jsonb), 'direction', k->'direction'),
            updated_at = clock_timestamp();
        n := n + 1;
    end loop;
    perform public.app_bump('kpis', me.username);
    return jsonb_build_object('saved', n);
end $$;

create or replace function public.save_target_year(p_token uuid, p_year int, p_data jsonb)
returns void language plpgsql security definer set search_path = public as $$
declare me public.app_users;
begin
    me := public.app_require(p_token, array['admin','manager']);
    insert into public.target_years(year, data, updated_at) values (p_year, p_data, clock_timestamp())
    on conflict (year) do update set data = excluded.data, updated_at = clock_timestamp();
    perform public.app_bump('targets', me.username);
end $$;

-- ---------- audit ----------
create or replace function public.add_audit(p_token uuid, p_entries jsonb)
returns void language plpgsql security definer set search_path = public as $$
declare me public.app_users;
begin
    me := public.app_session_user(p_token);
    insert into public.audit_log(username, display_name, action, item, old_value, new_value, reason)
    select me.username, me.display_name, e->>'action', e->>'item', e->>'oldValue', e->>'newValue', e->>'reason'
      from jsonb_array_elements(coalesce(p_entries, '[]'::jsonb)) e;
end $$;

create or replace function public.list_audit(p_token uuid, p_limit int default 1000)
returns jsonb language plpgsql security definer set search_path = public as $$
begin
    perform public.app_require(p_token, array['admin']);
    return coalesce((select jsonb_agg(jsonb_build_object('timestamp', at, 'username', username, 'displayName', display_name, 'action', action,
                     'item', item, 'oldValue', old_value, 'newValue', new_value, 'reason', reason) order by id desc)
                     from (select * from public.audit_log order by id desc limit least(coalesce(p_limit, 1000), 5000)) a), '[]'::jsonb);
end $$;

-- ---------- grants ----------
revoke all on public.app_users, public.app_sessions, public.audit_log from anon, authenticated;
grant select on public.kpis, public.records, public.target_years, public.app_events to anon, authenticated;
grant execute on function public.public_users(), public.login(text, text), public.logout(uuid), public.whoami(uuid),
    public.change_password(uuid, text, text), public.list_users(uuid), public.save_user(uuid, text, text, text, boolean, text),
    public.upsert_records(uuid, jsonb), public.delete_records(uuid, text[]), public.import_records(uuid, jsonb, text[]),
    public.save_kpis(uuid, jsonb, boolean), public.save_kpi_targets(uuid, jsonb), public.delete_kpi(uuid, text), public.save_target_year(uuid, int, jsonb),
    public.add_audit(uuid, jsonb), public.list_audit(uuid, int) to anon, authenticated;
revoke execute on function public.app_session_user(uuid), public.app_require(uuid, text[]), public.app_bump(text, text) from public, anon, authenticated;

-- ---------- realtime: dashboards listen to app_events ----------
do $$ begin
    if exists (select 1 from pg_publication where pubname = 'supabase_realtime')
       and not exists (select 1 from pg_publication_tables where pubname = 'supabase_realtime' and tablename = 'app_events') then
        alter publication supabase_realtime add table public.app_events;
    end if;
end $$;

-- ---------- first accounts (temporary passwords; each must be changed at first login) ----------
insert into public.app_users(username, display_name, role, pass_hash, must_change) values
    ('admin',   'عبدالعزيز عبدالله ال زربه', 'admin',   extensions.crypt('Admin@2026', extensions.gen_salt('bf')), true),
    ('manager', 'لمى علي الحقباني',          'manager', extensions.crypt('Manager@2026', extensions.gen_salt('bf')), true),
    ('data',    'وحدة مؤشرات الأداء',         'data',    extensions.crypt('Data@2026', extensions.gen_salt('bf')), true)
on conflict (username) do nothing;

-- ---------- make the new functions visible to the API right away ----------
-- PostgREST caches the schema; without a reload the page can get "could not find the function" (PGRST202).
notify pgrst, 'reload schema';
