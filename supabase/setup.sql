-- =====================================================================
-- لوحة مؤشرات الأداء — مستشفى إرادة والصحة النفسية بأبها
-- Supabase setup: run once in the Supabase SQL Editor (safe to re-run).
--
-- Model
--   * Excel import, master data and record history: see the "Official Excel template import" section
--     and docs/excel-import.md.
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
    -- Read by the record_history trigger (transaction-local).
    perform set_config('app.user', u.username, true);
    perform set_config('app.user_name', u.display_name, true);
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
-- Rules: data entry is part of the performance team, so its readings are approved on save and it may change
-- its own readings (approved or not); managers/admins may save, approve and overwrite any reading.
create or replace function public.upsert_records(p_token uuid, p_rows jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; r jsonb; ex public.records; approver boolean; saved int := 0; skipped int := 0; want_approved boolean; rid text;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    approver := me.role in ('admin','manager');
    for r in select * from jsonb_array_elements(coalesce(p_rows, '[]'::jsonb)) loop
        rid := (r->>'kpiCode') || '|' || (r->>'date');
        select * into ex from public.records where id = rid;
        if found and not ex.deleted and not approver and ex.created_user is distinct from me.username then
            skipped := skipped + 1; continue;
        end if;
        want_approved := (approver and coalesce((r->>'approved')::boolean, false)) or me.role = 'data';
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

-- =====================================================================
-- Official Excel template import (Excel → validation → mapping → database → KPI engine → dashboard)
--   * Master data: facilities, departments, categories (codes shared with the template).
--   * records gains the template's unique key and dimensions; one row per
--     Facility_ID + Indicator_ID + period start + Department_ID + Period_Type.
--     Rows of the main facility and the KPI's own department keep the id kpi_code|date,
--     so readings typed in the platform and imported from Excel are the same row.
--   * Every import is a batch (import_batches) with its rejected rows (import_errors).
--   * record_history keeps every insert / update / soft delete with old and new values.
--   * Readings are never hard-deleted.
-- =====================================================================

create table if not exists public.facilities (
    facility_id       text primary key check (facility_id ~ '^[A-Z0-9_]{2,20}$'),
    facility_name     text,
    facility_name_ar  text not null,
    cluster           text,
    region            text,
    is_primary        boolean not null default false,
    active            boolean not null default true,
    created_at        timestamptz not null default now(),
    updated_at        timestamptz not null default now()
);

create table if not exists public.departments (
    department_id         text primary key check (department_id ~ '^[A-Z0-9_]{2,20}$'),
    department_name       text,
    department_name_ar    text not null,
    parent_department_id  text references public.departments(department_id),
    department_type       text,
    active                boolean not null default true,
    created_at            timestamptz not null default now(),
    updated_at            timestamptz not null default now()
);

create table if not exists public.categories (
    category_id         text primary key check (category_id ~ '^[A-Z0-9_]{2,20}$'),
    category_name       text,
    category_name_ar    text not null,
    parent_category_id  text references public.categories(category_id),
    sort_order          int not null default 0,
    active              boolean not null default true,
    created_at          timestamptz not null default now(),
    updated_at          timestamptz not null default now()
);

-- BEGIN MASTER DATA (generated by tools/build_excel_template.py; same codes as the Excel template)
insert into public.facilities(facility_id, facility_name, facility_name_ar, cluster, region, active, is_primary) values
    ('FAC001', 'Eradah and Mental Health Hospital - Abha', 'مستشفى إرادة والصحة النفسية بأبها', 'تجمع عسير الصحي', 'عسير', true, true)
on conflict (facility_id) do nothing;
insert into public.departments(department_id, department_name, department_name_ar, parent_department_id, department_type, active) values
    ('DEP001', 'Medical Services', 'الخدمات العلاجية', null, 'إدارة', true),
    ('DEP002', 'Pharmacy / Medication Safety', 'الصيدلية / السلامة الدوائية', 'DEP001', 'قسم', true),
    ('DEP003', 'Support Services', 'الخدمات المساندة', null, 'إدارة', true),
    ('DEP004', 'Laboratory', 'المختبر', 'DEP001', 'قسم', true),
    ('DEP005', 'Emergency Department', 'الطوارئ', 'DEP001', 'قسم', true),
    ('DEP006', 'Outpatient Clinics', 'العيادات الخارجية', 'DEP001', 'قسم', true),
    ('DEP007', 'Medical Coordination', 'التنسيق الطبي', null, 'إدارة', true),
    ('DEP008', 'Quality Management', 'إدارة الجودة', null, 'إدارة', true),
    ('DEP009', 'Occupational Health', 'الصحة المهنية', 'DEP008', 'قسم', true),
    ('DEP010', 'Engineering Affairs', 'الشؤون الهندسية', null, 'إدارة', true),
    ('DEP011', '937 Call Center', 'مركز 937', null, 'وحدة', true),
    ('DEP012', 'Patient Experience', 'تجربة المريض', null, 'إدارة', true),
    ('DEP013', 'Finance', 'الإدارة المالية', null, 'إدارة', true),
    ('DEP014', 'Nursing Administration', 'إدارة التمريض', null, 'إدارة', true)
on conflict (department_id) do nothing;
insert into public.categories(category_id, category_name, category_name_ar, parent_category_id, sort_order) values
    ('CAT01', 'Medical Services', 'الخدمات العلاجية', null, 0),
    ('CAT02', 'Pharmacy', 'الصيدلية', 'CAT01', 1),
    ('CAT03', 'Laboratory (Enayati)', 'المختبر (عينتي)', 'CAT01', 2),
    ('CAT04', 'Emergency', 'الطوارئ', null, 3),
    ('CAT05', 'Clinics', 'العيادات', null, 4),
    ('CAT06', 'Quality', 'الجودة', null, 5),
    ('CAT07', 'Engineering Affairs', 'الشؤون الهندسية', null, 6),
    ('CAT08', 'Medical Coordination', 'التنسيق الطبي', null, 7),
    ('CAT09', '937 Reports', 'بلاغات 937', null, 8),
    ('CAT10', 'Patient Experience', 'تجربة المريض', null, 9),
    ('CAT11', 'Self Revenue', 'الموارد الذاتية', null, 10),
    ('CAT12', 'Nursing', 'التمريض', null, 11)
on conflict (category_id) do nothing;
-- END MASTER DATA

alter table public.records add column if not exists record_key      text;
alter table public.records add column if not exists facility_id     text;
alter table public.records add column if not exists department_id   text;
alter table public.records add column if not exists period_type     text;
alter table public.records add column if not exists period_end      date;
alter table public.records add column if not exists target_value    double precision;
alter table public.records add column if not exists unit            text;
alter table public.records add column if not exists data_entry_by   text;
alter table public.records add column if not exists data_entry_date date;
alter table public.records add column if not exists batch_id        bigint;
alter table public.records add column if not exists source_row      int;

create sequence if not exists public.import_batch_seq;
create table if not exists public.import_batches (
    batch_id            bigint primary key default nextval('public.import_batch_seq'),
    batch_code          text unique,
    file_name           text not null,
    file_size           bigint,
    file_hash           text,
    sheet_name          text,
    template_version    text,
    uploaded_by         text not null,
    uploaded_by_name    text,
    upload_date         timestamptz not null default now(),
    finished_at         timestamptz,
    records_received    int not null default 0,
    records_empty       int not null default 0,
    records_duplicates  int not null default 0,
    records_inserted    int not null default 0,
    records_updated     int not null default 0,
    records_unchanged   int not null default 0,
    records_rejected    int not null default 0,
    records_warnings    int not null default 0,
    approve             boolean not null default false,
    status              text not null default 'IN_PROGRESS'
                        check (status in ('IN_PROGRESS','COMPLETED','COMPLETED_WITH_WARNINGS','COMPLETED_WITH_ERRORS','FAILED','CANCELLED')),
    notes               text
);
create index if not exists import_batches_date_idx on public.import_batches (upload_date desc);

create table if not exists public.import_errors (
    error_id        bigserial primary key,
    batch_id        bigint not null references public.import_batches(batch_id),
    row_number      int,
    record_key      text,
    indicator_id    text,
    field_name      text,
    error_code      text,
    error_type      text check (error_type in ('ERROR','WARNING','DUPLICATE')),
    error_message   text,
    suggested_fix   text,
    original_value  text,
    created_at      timestamptz not null default now()
);
create index if not exists import_errors_batch_idx on public.import_errors (batch_id, row_number);

create table if not exists public.record_history (
    history_id       bigserial primary key,
    record_id        text not null,
    record_key       text,
    kpi_code         text,
    batch_id         bigint,
    action           text not null check (action in ('INSERT','UPDATE','APPROVE','SOFT_DELETE','RESTORE')),
    changed_fields   text[],
    old_values       jsonb,
    new_values       jsonb,
    changed_by       text,
    changed_by_name  text,
    changed_at       timestamptz not null default now(),
    reason           text
);
create index if not exists record_history_record_idx on public.record_history (record_id, changed_at desc);
create index if not exists record_history_at_idx on public.record_history (changed_at desc);
create index if not exists record_history_batch_idx on public.record_history (batch_id);

alter table public.facilities     enable row level security;
alter table public.departments    enable row level security;
alter table public.categories     enable row level security;
alter table public.import_batches enable row level security;
alter table public.import_errors  enable row level security;
alter table public.record_history enable row level security;
drop policy if exists facilities_read on public.facilities;
create policy facilities_read on public.facilities for select using (true);
drop policy if exists departments_read on public.departments;
create policy departments_read on public.departments for select using (true);
drop policy if exists categories_read on public.categories;
create policy categories_read on public.categories for select using (true);

-- ---------- import helpers ----------
-- First date the platform accepts. There is no end date: any date up to today (Riyadh) is valid.
create or replace function public.app_data_start()
returns date language sql immutable as $$ select date '2026-01-01' $$;

create or replace function public.app_today()
returns date language sql stable as $$ select (now() at time zone 'Asia/Riyadh')::date $$;

create or replace function public.app_period_type(p_frequency text)
returns text language sql immutable as $$
    select case lower(coalesce(p_frequency, '')) when 'weekly' then 'WEEKLY' when 'monthly' then 'MONTHLY'
                when 'quarterly' then 'QUARTERLY' when 'annual' then 'ANNUAL' else 'DAILY' end
$$;

-- Weeks start on Sunday (as in the platform); months and quarters on their first day.
create or replace function public.app_period_start(p_date date, p_period text)
returns date language sql immutable as $$
    select case p_period when 'WEEKLY' then p_date - extract(dow from p_date)::int
                         when 'MONTHLY' then date_trunc('month', p_date)::date
                         when 'QUARTERLY' then date_trunc('quarter', p_date)::date
                         when 'ANNUAL' then date_trunc('year', p_date)::date else p_date end
$$;

create or replace function public.app_period_end(p_start date, p_period text)
returns date language sql immutable as $$
    select case p_period when 'WEEKLY' then p_start + 6
                         when 'MONTHLY' then (date_trunc('month', p_start) + interval '1 month - 1 day')::date
                         when 'QUARTERLY' then (date_trunc('quarter', p_start) + interval '3 months - 1 day')::date
                         when 'ANNUAL' then (date_trunc('year', p_start) + interval '1 year - 1 day')::date else p_start end
$$;

create or replace function public.app_record_key(p_facility text, p_kpi text, p_start date, p_department text, p_period text)
returns text language sql immutable as $$
    select upper(p_facility || '-' || p_kpi || '-' || to_char(p_start, 'YYYYMMDD') || '-' || p_department || '-' || p_period)
$$;

create or replace function public.app_primary_facility()
returns text language sql stable security definer set search_path = public as $$
    select coalesce((select facility_id from public.facilities where is_primary and active order by facility_id limit 1), 'FAC001')
$$;

-- The KPI's own department: its departmentId, else the department with its Arabic name.
create or replace function public.app_kpi_department(p_kpi jsonb)
returns text language sql stable security definer set search_path = public as $$
    select coalesce(nullif(p_kpi->>'departmentId', ''),
                    (select department_id from public.departments where department_name_ar = p_kpi->>'department' order by department_id limit 1),
                    'UNASSIGNED')
$$;

-- Calculation type of a KPI: its calculationType, else inferred from the aggregation.
create or replace function public.app_calc_type(p_kpi jsonb)
returns text language sql immutable as $$
    select coalesce(nullif(p_kpi->>'calculationType', ''),
                    case p_kpi->>'aggregation' when 'RATIO' then 'PERCENTAGE'
                         when 'SUM' then case when coalesce(p_kpi->>'unit', 'عدد') = 'عدد' then 'COUNT' else 'SUM' end
                         when 'WEIGHTED_AVERAGE' then 'WEIGHTED_AVERAGE' when 'LAST_VALUE' then 'LAST_VALUE' else 'AVERAGE' end)
$$;

-- Numerator ÷ denominator is multiplied by this; 0 means the value is entered directly.
create or replace function public.app_calc_multiplier(p_calc text)
returns double precision language sql immutable as $$
    select case p_calc when 'PERCENTAGE' then 100 when 'COMPLIANCE' then 100 when 'RATE' then 100 when 'RATIO' then 1
                       when 'RATE_PER_1000' then 1000 when 'RATE_PER_100000' then 100000 else 0 end::double precision
$$;

-- Numbers from an untrusted file: a JSON number, or a plain numeric string. Anything else is rejected.
create or replace function public.app_num(p jsonb, out val double precision, out ok boolean)
language plpgsql immutable as $$
declare s text;
begin
    ok := true; val := null;
    if p is null or jsonb_typeof(p) = 'null' then return; end if;
    s := btrim(p #>> '{}');
    if jsonb_typeof(p) = 'string' and s = '' then return; end if;
    if jsonb_typeof(p) not in ('number', 'string') or s !~ '^-?[0-9]+(\.[0-9]+)?([eE][-+]?[0-9]+)?$' or length(s) > 40 then ok := false; return; end if;
    val := s::double precision;
    if val in ('Infinity'::float8, '-Infinity'::float8) then ok := false; val := null; end if;
end $$;

-- ---------- records: keys, history, no hard delete ----------
create or replace function public.records_fill_keys()
returns trigger language plpgsql security definer set search_path = public as $$
declare k jsonb;
begin
    if new.facility_id is null or new.department_id is null or new.period_type is null or new.unit is null then
        select data into k from public.kpis where code = new.kpi_code;
        new.facility_id := coalesce(new.facility_id, public.app_primary_facility());
        new.department_id := coalesce(new.department_id, public.app_kpi_department(coalesce(k, '{}'::jsonb)));
        new.period_type := coalesce(new.period_type, public.app_period_type(k->>'frequency'));
        new.unit := coalesce(new.unit, k->>'unit');
    end if;
    new.period_end := public.app_period_end(new.date, new.period_type);
    new.record_key := public.app_record_key(new.facility_id, new.kpi_code, new.date, new.department_id, new.period_type);
    return new;
end $$;
drop trigger if exists records_fill_keys on public.records;
create trigger records_fill_keys before insert or update on public.records for each row execute function public.records_fill_keys();

-- Existing readings get their key once (before the history trigger exists, so this is not logged as edits).
update public.records set record_key = null where record_key is null;
create unique index if not exists records_record_key_uidx on public.records (record_key);
create index if not exists records_batch_idx on public.records (batch_id);

create or replace function public.app_record_values(r public.records)
returns jsonb language sql immutable as $$
    select jsonb_build_object('date', r.date, 'actual', r.actual, 'numerator', r.numerator, 'denominator', r.denominator,
        'weight', r.weight, 'data_source', r.data_source, 'notes', r.notes, 'approved', r.approved, 'deleted', r.deleted,
        'facility_id', r.facility_id, 'department_id', r.department_id, 'period_type', r.period_type)
$$;

-- Who and why come from the calling function (app_require sets app.user; imports set app.batch_id and app.reason).
create or replace function public.records_history()
returns trigger language plpgsql security definer set search_path = public as $$
declare o jsonb; n jsonb; changed text[]; act text;
begin
    n := public.app_record_values(new);
    if tg_op = 'UPDATE' then
        o := public.app_record_values(old);
        select array_agg(e.key order by e.key) into changed from jsonb_each(n) e where e.value is distinct from o->e.key;
        if changed is null then return null; end if;
        act := case when new.deleted and not old.deleted then 'SOFT_DELETE'
                    when old.deleted and not new.deleted then 'RESTORE'
                    when changed = array['approved'] and new.approved then 'APPROVE' else 'UPDATE' end;
    else
        act := 'INSERT';
    end if;
    insert into public.record_history(record_id, record_key, kpi_code, batch_id, action, changed_fields, old_values, new_values,
                                      changed_by, changed_by_name, reason)
    values (new.id, new.record_key, new.kpi_code, nullif(current_setting('app.batch_id', true), '')::bigint, act, changed, o, n,
            coalesce(nullif(current_setting('app.user', true), ''), new.updated_by, new.created_by),
            coalesce(nullif(current_setting('app.user_name', true), ''), new.updated_by, new.created_by),
            nullif(current_setting('app.reason', true), ''));
    return null;
end $$;
drop trigger if exists records_history on public.records;
create trigger records_history after insert or update on public.records for each row execute function public.records_history();

create or replace function public.app_forbid_delete()
returns trigger language plpgsql as $$
begin
    raise exception 'HARD_DELETE_FORBIDDEN: % rows are kept; use soft delete', tg_table_name using errcode = '42501';
end $$;
drop trigger if exists records_no_delete on public.records;
create trigger records_no_delete before delete on public.records for each row execute function public.app_forbid_delete();
drop trigger if exists record_history_append_only on public.record_history;
create trigger record_history_append_only before update or delete on public.record_history for each row execute function public.app_forbid_delete();
drop trigger if exists import_errors_append_only on public.import_errors;
create trigger import_errors_append_only before update or delete on public.import_errors for each row execute function public.app_forbid_delete();

-- ---------- import API ----------
-- Step 1: open a batch. p_meta: fileName, fileSize, fileHash, sheetName, templateVersion, received, empty, approve.
create or replace function public.import_start(p_token uuid, p_meta jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; b bigint; code text;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    -- A batch left open (closed browser, lost connection) is closed as failed.
    update public.import_batches set status = 'FAILED', finished_at = now(), notes = coalesce(notes, 'لم يكتمل الرفع')
     where uploaded_by = me.username and status = 'IN_PROGRESS' and upload_date < now() - interval '1 hour';
    insert into public.import_batches(file_name, file_size, file_hash, sheet_name, template_version, uploaded_by, uploaded_by_name,
                                      records_received, records_empty, approve)
    values (left(coalesce(nullif(btrim(p_meta->>'fileName'), ''), 'file.xlsx'), 255),
            (public.app_num(p_meta->'fileSize')).val::bigint, left(p_meta->>'fileHash', 128), left(p_meta->>'sheetName', 64),
            left(p_meta->>'templateVersion', 20), me.username, me.display_name,
            coalesce((public.app_num(p_meta->'received')).val::int, 0), coalesce((public.app_num(p_meta->'empty')).val::int, 0),
            (me.role in ('admin','manager') and coalesce((p_meta->>'approve')::boolean, true)) or me.role = 'data')
    returning batch_id into b;
    code := to_char(public.app_today(), 'YYYY') || '-' || lpad(b::text, 5, '0');
    update public.import_batches set batch_code = code where batch_id = b;
    return jsonb_build_object('batchId', b, 'batchCode', code);
end $$;

-- Step 2 (repeated, ≤ 1000 rows per call): every row is checked again here; the file is untrusted input.
-- Row: rowNumber, facilityId, departmentId, indicatorId, periodType, measurementDate (YYYY-MM-DD), numerator,
-- denominator, actualValue, dataSource, dataEntryBy, dataEntryDate, notes.
-- Returns counts and the rejected rows [{row, code, field, value, key}].
create or replace function public.import_rows(p_token uuid, p_batch_id bigint, p_rows jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare
    me public.app_users; b public.import_batches; r jsonb; k jsonb; kmap jsonb; ex public.records; found_ex boolean;
    approver boolean; approve boolean; prim text; today date; errs jsonb := '[]'::jsonb; row_errs jsonb;
    rn int; fac text; dep text; ind text; per text; d date; ps date; key text; rid text; home text; calc text; mult double precision;
    num double precision; den double precision; act double precision; val double precision; vmin double precision; vmax double precision;
    okn boolean; okd boolean; oka boolean; src text; note text; tv double precision; entry_date date; direct boolean;
    ins int := 0; upd int := 0; same int := 0; rej int := 0;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    select * into b from public.import_batches where batch_id = p_batch_id for update;
    if not found or b.uploaded_by <> me.username or b.status <> 'IN_PROGRESS' then raise exception 'BAD_BATCH' using errcode = '22023'; end if;
    if jsonb_typeof(p_rows) <> 'array' or jsonb_array_length(p_rows) > 1000 then raise exception 'TOO_MANY_ROWS' using errcode = '22023'; end if;
    approver := me.role in ('admin','manager');
    approve := b.approve;
    perform set_config('app.batch_id', p_batch_id::text, true);
    perform set_config('app.reason', left('استيراد Excel ' || coalesce(b.batch_code, '') || ': ' || b.file_name, 300), true);
    select coalesce(jsonb_object_agg(code, data), '{}'::jsonb) into kmap from public.kpis;
    prim := public.app_primary_facility();
    today := public.app_today();

    for r in select * from jsonb_array_elements(p_rows) loop
        row_errs := '[]'::jsonb;
        rn := (public.app_num(r->'rowNumber')).val::int;
        fac := upper(btrim(coalesce(r->>'facilityId', '')));
        dep := upper(btrim(coalesce(r->>'departmentId', '')));
        ind := upper(btrim(coalesce(r->>'indicatorId', '')));
        per := upper(btrim(coalesce(r->>'periodType', '')));
        k := case when ind = '' then null else kmap->ind end;
        d := null;
        if fac = '' or dep = '' or ind = '' or per = '' or coalesce(r->>'measurementDate', '') = '' then
            row_errs := row_errs || jsonb_build_object('code', 'E11', 'field', 'Unique_Key');
        end if;
        if ind <> '' and k is null then row_errs := row_errs || jsonb_build_object('code', 'E04', 'field', 'Indicator_ID', 'value', ind);
        elsif k is not null and (k->>'active') = 'false' then row_errs := row_errs || jsonb_build_object('code', 'E13', 'field', 'Indicator_ID', 'value', ind);
        end if;
        if fac <> '' and not exists (select 1 from public.facilities where facility_id = fac and active) then
            row_errs := row_errs || jsonb_build_object('code', 'E06', 'field', 'Facility_ID', 'value', fac);
        end if;
        if dep <> '' and not exists (select 1 from public.departments where department_id = dep and active) then
            row_errs := row_errs || jsonb_build_object('code', 'E05', 'field', 'Department_ID', 'value', dep);
        end if;
        if per <> '' and (per not in ('DAILY','WEEKLY','MONTHLY','QUARTERLY','ANNUAL') or (k is not null and per <> public.app_period_type(k->>'frequency'))) then
            row_errs := row_errs || jsonb_build_object('code', 'E12', 'field', 'Period_Type', 'value', per);
        end if;
        if coalesce(r->>'measurementDate', '') <> '' then
            if (r->>'measurementDate') ~ '^\d{4}-\d{2}-\d{2}$' then
                begin d := (r->>'measurementDate')::date; exception when others then d := null; end;
            end if;
            if d is null then row_errs := row_errs || jsonb_build_object('code', 'E02', 'field', 'Measurement_Date', 'value', left(r->>'measurementDate', 40));
            elsif d < public.app_data_start() or d > today then row_errs := row_errs || jsonb_build_object('code', 'E03', 'field', 'Measurement_Date', 'value', d::text);
            end if;
        end if;
        select x.val, x.ok into num, okn from public.app_num(r->'numerator') x;
        select x.val, x.ok into den, okd from public.app_num(r->'denominator') x;
        select x.val, x.ok into act, oka from public.app_num(r->'actualValue') x;
        if not okn then row_errs := row_errs || jsonb_build_object('code', 'E07', 'field', 'Numerator', 'value', left(r->>'numerator', 40)); end if;
        if not okd then row_errs := row_errs || jsonb_build_object('code', 'E07', 'field', 'Denominator', 'value', left(r->>'denominator', 40)); end if;
        if not oka then row_errs := row_errs || jsonb_build_object('code', 'E07', 'field', 'Actual_Value', 'value', left(r->>'actualValue', 40)); end if;
        val := null; direct := false;
        if k is not null and okn and okd and oka then
            calc := public.app_calc_type(k);
            mult := public.app_calc_multiplier(calc);
            if mult > 0 then
                -- A rate entered as its final number (no numerator / denominator) is taken as is.
                if num is null and den is null and act is not null then val := act; direct := true;
                elsif num is null or den is null then row_errs := row_errs || jsonb_build_object('code', 'E01', 'field', case when num is null then 'Numerator' else 'Denominator' end);
                elsif den = 0 then row_errs := row_errs || jsonb_build_object('code', 'E08', 'field', 'Denominator', 'value', '0');
                elsif num < 0 or den < 0 then row_errs := row_errs || jsonb_build_object('code', 'E09', 'field', case when num < 0 then 'Numerator' else 'Denominator' end);
                else val := num / den * mult;
                end if;
            elsif act is null then row_errs := row_errs || jsonb_build_object('code', 'E01', 'field', 'Actual_Value');
            else val := act;
            end if;
            -- Allowed range: validMin / validMax on the KPI; by default no negatives and percentages up to 100.
            vmin := case when k ? 'validMin' then (public.app_num(k->'validMin')).val else 0 end;
            vmax := case when k ? 'validMax' then (public.app_num(k->'validMax')).val when k->>'unit' = '%' and calc not in ('PERCENTAGE') then 100 end;
            -- A final rate above the maximum is imported as entered (the platform warns about it); below the minimum is rejected.
            if val is not null and ((vmin is not null and val < vmin) or (not direct and vmax is not null and val > vmax)) then
                row_errs := row_errs || jsonb_build_object('code', 'E09', 'field', case when mult > 0 and not direct then 'Numerator' else 'Actual_Value' end, 'value', round(val::numeric, 4)::text);
            end if;
        end if;

        ps := case when d is not null and per in ('DAILY','WEEKLY','MONTHLY','QUARTERLY','ANNUAL') then public.app_period_start(d, per) end;
        key := case when ps is not null and fac <> '' and dep <> '' and ind <> '' then public.app_record_key(fac, ind, ps, dep, per) end;
        if jsonb_array_length(row_errs) > 0 then
            rej := rej + 1;
            errs := errs || (select jsonb_agg(e || jsonb_build_object('row', rn, 'key', key, 'indicator', ind)) from jsonb_array_elements(row_errs) e);
            continue;
        end if;

        home := public.app_kpi_department(k);
        rid := case when fac = prim and dep = home then ind || '|' || ps::text else key end;
        select * into ex from public.records where record_key = key;
        found_ex := found;
        if not found_ex then select * into ex from public.records where id = rid; found_ex := found; end if;
        src := left(nullif(btrim(r->>'dataSource'), ''), 200);
        note := left(nullif(btrim(r->>'notes'), ''), 1000);
        -- Data entry may only change its own readings.
        if found_ex and not ex.deleted and not approver and ex.created_user is distinct from me.username then
            rej := rej + 1;
            errs := errs || jsonb_build_object('row', rn, 'key', key, 'indicator', ind, 'code', 'E16', 'field', 'Unique_Key');
            continue;
        end if;
        if found_ex and not ex.deleted and ex.actual is not distinct from val and ex.quarter is not distinct from ('Q' || extract(quarter from d)::int) and ex.numerator is not distinct from (case when mult > 0 then num end)
           and ex.denominator is not distinct from (case when mult > 0 then den end)
           and coalesce(ex.notes, '') = coalesce(note, '') and coalesce(ex.data_source, '') = coalesce(src, '') and (ex.approved or not approve) then
            same := same + 1;
            continue;
        end if;
        tv := coalesce((public.app_num(k->'targets'->(extract(year from ps)::int::text)->'value')).val, (public.app_num(k->'target')).val);
        entry_date := case when coalesce(r->>'dataEntryDate', '') ~ '^\d{4}-\d{2}-\d{2}$' then (r->>'dataEntryDate')::date end;
        insert into public.records as t (id, kpi_code, date, year, quarter, actual, numerator, denominator, data_source, notes,
                                         approved, created_by, created_user, created_at, updated_by, updated_at, approved_by, approved_at, deleted,
                                         record_key, facility_id, department_id, period_type, target_value, unit, data_entry_by, data_entry_date,
                                         batch_id, source_row)
        -- The quarter follows the date the reading was entered on (a week can straddle two quarters).
        values (rid, ind, ps, extract(year from d)::int, 'Q' || extract(quarter from d)::int, val,
                case when mult > 0 then num end, case when mult > 0 then den end, src, note,
                approve, me.display_name, me.username, now(), me.display_name, clock_timestamp(),
                case when approve then me.display_name end, case when approve then now() end, false,
                key, fac, dep, per, tv, k->>'unit', left(nullif(btrim(r->>'dataEntryBy'), ''), 120), entry_date, p_batch_id, rn)
        on conflict (id) do update set
            actual = excluded.actual, numerator = excluded.numerator, denominator = excluded.denominator,
            year = excluded.year, quarter = excluded.quarter, data_source = excluded.data_source, notes = excluded.notes,
            approved = case when excluded.approved then true when t.deleted then false else t.approved and t.actual is not distinct from excluded.actual end,
            created_by = case when t.deleted then excluded.created_by else t.created_by end,
            created_user = case when t.deleted then excluded.created_user else t.created_user end,
            updated_by = me.display_name, updated_at = clock_timestamp(), deleted = false, rev = nextval('public.records_rev_seq'),
            approved_by = case when excluded.approved and not t.approved then me.display_name when excluded.approved then t.approved_by end,
            approved_at = case when excluded.approved and not t.approved then now() when excluded.approved then t.approved_at end,
            facility_id = excluded.facility_id, department_id = excluded.department_id, period_type = excluded.period_type,
            target_value = excluded.target_value, unit = excluded.unit, data_entry_by = excluded.data_entry_by,
            data_entry_date = excluded.data_entry_date, batch_id = excluded.batch_id, source_row = excluded.source_row;
        if found_ex and not ex.deleted then upd := upd + 1; else ins := ins + 1; end if;
    end loop;

    update public.import_batches set records_inserted = records_inserted + ins, records_updated = records_updated + upd,
           records_unchanged = records_unchanged + same, records_rejected = records_rejected + rej
     where batch_id = p_batch_id;
    return jsonb_build_object('inserted', ins, 'updated', upd, 'unchanged', same, 'rejected', rej, 'errors', errs);
end $$;

-- Step 3: close the batch with the row report (rows rejected before upload, duplicates, warnings and the
-- server's rejections) and tell the open dashboards to refresh.
-- p_summary: duplicates, warnings, rejectedBeforeUpload, cancelled.
create or replace function public.import_finish(p_token uuid, p_batch_id bigint, p_errors jsonb, p_summary jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; b public.import_batches; st text;
begin
    me := public.app_require(p_token, array['admin','manager','data']);
    select * into b from public.import_batches where batch_id = p_batch_id for update;
    if not found or b.uploaded_by <> me.username or b.status <> 'IN_PROGRESS' then raise exception 'BAD_BATCH' using errcode = '22023'; end if;
    insert into public.import_errors(batch_id, row_number, record_key, indicator_id, field_name, error_code, error_type, error_message, suggested_fix, original_value)
    select p_batch_id, (public.app_num(e->'row')).val::int, left(e->>'key', 120), left(e->>'indicator', 64), left(e->>'field', 64), left(e->>'code', 8),
           case when e->>'code' like 'E%' then 'ERROR' when e->>'code' = 'W05' then 'DUPLICATE' else 'WARNING' end,
           left(e->>'message', 500), left(e->>'fix', 500), left(e->>'value', 200)
      from (select e from jsonb_array_elements(coalesce(p_errors, '[]'::jsonb)) e limit 50000) x;
    update public.import_batches set
        records_duplicates = coalesce((public.app_num(p_summary->'duplicates')).val::int, 0),
        records_warnings = coalesce((public.app_num(p_summary->'warnings')).val::int, 0),
        records_rejected = records_rejected + coalesce((public.app_num(p_summary->'rejectedBeforeUpload')).val::int, 0),
        finished_at = now()
     where batch_id = p_batch_id returning * into b;
    st := case when coalesce((p_summary->>'cancelled')::boolean, false) then 'CANCELLED'
               when b.records_rejected > 0 then 'COMPLETED_WITH_ERRORS'
               when b.records_warnings > 0 or b.records_duplicates > 0 then 'COMPLETED_WITH_WARNINGS' else 'COMPLETED' end;
    update public.import_batches set status = st where batch_id = p_batch_id;
    insert into public.audit_log(username, display_name, action, item, new_value, reason)
    values (me.username, me.display_name, 'IMPORT_TEMPLATE', b.batch_code || ' — ' || b.file_name,
            format('%s سجل: %s جديد، %s محدّث، %s دون تغيير، %s مرفوض', b.records_received, b.records_inserted, b.records_updated, b.records_unchanged, b.records_rejected),
            'استيراد قالب Excel الرسمي');
    perform public.app_bump('records', me.username);
    return jsonb_build_object('batchId', b.batch_id, 'batchCode', b.batch_code, 'status', st, 'received', b.records_received,
        'inserted', b.records_inserted, 'updated', b.records_updated, 'unchanged', b.records_unchanged, 'rejected', b.records_rejected,
        'duplicates', b.records_duplicates, 'warnings', b.records_warnings, 'empty', b.records_empty);
end $$;

create or replace function public.app_batch_json(b public.import_batches)
returns jsonb language sql immutable as $$
    select jsonb_build_object('batchId', b.batch_id, 'batchCode', b.batch_code, 'fileName', b.file_name, 'fileSize', b.file_size,
        'fileHash', b.file_hash, 'sheetName', b.sheet_name, 'templateVersion', b.template_version, 'uploadedBy', b.uploaded_by,
        'uploadedByName', b.uploaded_by_name, 'uploadDate', b.upload_date, 'finishedAt', b.finished_at, 'received', b.records_received,
        'empty', b.records_empty, 'duplicates', b.records_duplicates, 'inserted', b.records_inserted, 'updated', b.records_updated,
        'unchanged', b.records_unchanged, 'rejected', b.records_rejected, 'warnings', b.records_warnings, 'approve', b.approve,
        'status', b.status, 'notes', b.notes)
$$;

create or replace function public.list_import_batches(p_token uuid, p_limit int default 200)
returns jsonb language plpgsql security definer set search_path = public as $$
begin
    perform public.app_require(p_token, array['admin','manager','data']);
    return coalesce((select jsonb_agg(public.app_batch_json(b) order by b.batch_id desc)
                       from (select * from public.import_batches order by batch_id desc limit least(coalesce(p_limit, 200), 1000)) b), '[]'::jsonb);
end $$;

create or replace function public.get_import_batch(p_token uuid, p_batch_id bigint)
returns jsonb language plpgsql security definer set search_path = public as $$
declare b public.import_batches;
begin
    perform public.app_require(p_token, array['admin','manager','data']);
    select * into b from public.import_batches where batch_id = p_batch_id;
    if not found then raise exception 'BAD_BATCH' using errcode = '22023'; end if;
    return public.app_batch_json(b) || jsonb_build_object(
        'errors', coalesce((select jsonb_agg(jsonb_build_object('row', e.row_number, 'key', e.record_key, 'indicator', e.indicator_id,
                     'field', e.field_name, 'code', e.error_code, 'type', e.error_type, 'message', e.error_message, 'fix', e.suggested_fix,
                     'value', e.original_value) order by e.row_number, e.error_id)
                   from (select * from public.import_errors where batch_id = p_batch_id order by row_number, error_id limit 20000) e), '[]'::jsonb),
        'history', (select count(*) from public.record_history where batch_id = p_batch_id));
end $$;

-- Change history of readings (admin and performance manager). p_filter: recordId, kpiCode, batchId, since (date).
create or replace function public.list_record_history(p_token uuid, p_filter jsonb default '{}'::jsonb, p_limit int default 500)
returns jsonb language plpgsql security definer set search_path = public as $$
begin
    perform public.app_require(p_token, array['admin','manager']);
    return coalesce((select jsonb_agg(jsonb_build_object('id', h.history_id, 'recordId', h.record_id, 'recordKey', h.record_key,
                'kpiCode', h.kpi_code, 'batchId', h.batch_id, 'action', h.action, 'fields', h.changed_fields, 'old', h.old_values,
                'new', h.new_values, 'by', h.changed_by, 'byName', h.changed_by_name, 'at', h.changed_at, 'reason', h.reason) order by h.history_id desc)
        from (select * from public.record_history
               where (p_filter->>'recordId' is null or record_id = p_filter->>'recordId')
                 and (p_filter->>'kpiCode' is null or kpi_code = p_filter->>'kpiCode')
                 and (p_filter->>'batchId' is null or batch_id = (p_filter->>'batchId')::bigint)
                 and (p_filter->>'since' is null or changed_at >= (p_filter->>'since')::date)
               order by history_id desc limit least(coalesce(p_limit, 500), 5000)) h), '[]'::jsonb);
end $$;

-- Master data (admin). p_kind: facilities | departments | categories. Rows are added or updated, never deleted
-- (set active = false instead), so readings keep their references.
create or replace function public.save_master_data(p_token uuid, p_kind text, p_rows jsonb)
returns jsonb language plpgsql security definer set search_path = public as $$
declare me public.app_users; r jsonb; n int := 0; id text;
begin
    me := public.app_require(p_token, array['admin']);
    for r in select * from jsonb_array_elements(coalesce(p_rows, '[]'::jsonb)) loop
        id := upper(btrim(coalesce(r->>'id', '')));
        if coalesce(btrim(r->>'nameAr'), '') = '' then raise exception 'NAME_REQUIRED' using errcode = '22023'; end if;
        if p_kind = 'facilities' then
            insert into public.facilities(facility_id, facility_name, facility_name_ar, cluster, region, is_primary, active)
            values (id, r->>'name', btrim(r->>'nameAr'), r->>'cluster', r->>'region', coalesce((r->>'primary')::boolean, false), coalesce((r->>'active')::boolean, true))
            on conflict (facility_id) do update set facility_name = excluded.facility_name, facility_name_ar = excluded.facility_name_ar,
                cluster = excluded.cluster, region = excluded.region, is_primary = excluded.is_primary, active = excluded.active, updated_at = now();
        elsif p_kind = 'departments' then
            insert into public.departments(department_id, department_name, department_name_ar, parent_department_id, department_type, active)
            values (id, r->>'name', btrim(r->>'nameAr'), nullif(upper(btrim(r->>'parentId')), ''), r->>'type', coalesce((r->>'active')::boolean, true))
            on conflict (department_id) do update set department_name = excluded.department_name, department_name_ar = excluded.department_name_ar,
                parent_department_id = excluded.parent_department_id, department_type = excluded.department_type, active = excluded.active, updated_at = now();
        elsif p_kind = 'categories' then
            insert into public.categories(category_id, category_name, category_name_ar, parent_category_id, sort_order, active)
            values (id, r->>'name', btrim(r->>'nameAr'), nullif(upper(btrim(r->>'parentId')), ''), coalesce((public.app_num(r->'sortOrder')).val::int, 0), coalesce((r->>'active')::boolean, true))
            on conflict (category_id) do update set category_name = excluded.category_name, category_name_ar = excluded.category_name_ar,
                parent_category_id = excluded.parent_category_id, sort_order = excluded.sort_order, active = excluded.active, updated_at = now();
        else
            raise exception 'BAD_KIND' using errcode = '22023';
        end if;
        n := n + 1;
    end loop;
    if p_kind = 'facilities' and (select count(*) from public.facilities where is_primary) > 1 then raise exception 'ONE_PRIMARY_FACILITY' using errcode = '22023'; end if;
    insert into public.audit_log(username, display_name, action, item, new_value, reason)
    values (me.username, me.display_name, 'EDIT_MASTER', p_kind, n::text, 'تحديث البيانات المرجعية');
    perform public.app_bump('master', me.username);
    return jsonb_build_object('saved', n);
end $$;

-- Readable views in the shape of the design (indicators = KPI master, indicator_data = readings).
create or replace view public.indicators with (security_invoker = true) as
    select k.code as indicator_id, k.data->>'name' as indicator_name, coalesce(k.data->>'categoryId', c.category_id) as category_id,
           k.data->>'category' as category_name, public.app_kpi_department(k.data) as department_id, k.data->>'unit' as unit,
           k.data->>'frequency' as frequency, public.app_calc_type(k.data) as calculation_type, k.data->>'direction' as direction,
           (public.app_num(k.data->'target')).val as target, coalesce((k.data->>'active')::boolean, true) as active, k.updated_at
      from public.kpis k left join public.categories c on c.category_name_ar = k.data->>'category';

create or replace view public.indicator_data with (security_invoker = true) as
    select r.id as record_id, r.record_key, r.kpi_code as indicator_id, r.department_id, r.facility_id, r.date as measurement_date,
           r.period_type, r.date as period_start, r.period_end, r.numerator, r.denominator, r.actual as actual_value, r.target_value,
           r.unit, r.data_source, r.notes, r.approved, r.batch_id, r.created_at, r.updated_at
      from public.records r where not r.deleted;

-- ---------- grants ----------
-- Some new projects do not give the API roles access to the public schema by default.
grant usage on schema public to anon, authenticated;
revoke all on public.app_users, public.app_sessions, public.audit_log, public.import_batches, public.import_errors, public.record_history from anon, authenticated;
grant select on public.kpis, public.records, public.target_years, public.app_events, public.facilities, public.departments, public.categories,
    public.indicators, public.indicator_data to anon, authenticated;
grant execute on function public.public_users(), public.login(text, text), public.logout(uuid), public.whoami(uuid),
    public.change_password(uuid, text, text), public.list_users(uuid), public.save_user(uuid, text, text, text, boolean, text),
    public.upsert_records(uuid, jsonb), public.delete_records(uuid, text[]), public.import_records(uuid, jsonb, text[]),
    public.save_kpis(uuid, jsonb, boolean), public.save_kpi_targets(uuid, jsonb), public.delete_kpi(uuid, text), public.save_target_year(uuid, int, jsonb),
    public.add_audit(uuid, jsonb), public.list_audit(uuid, int),
    public.import_start(uuid, jsonb), public.import_rows(uuid, bigint, jsonb), public.import_finish(uuid, bigint, jsonb, jsonb),
    public.list_import_batches(uuid, int), public.get_import_batch(uuid, bigint), public.list_record_history(uuid, jsonb, int),
    public.save_master_data(uuid, text, jsonb) to anon, authenticated;
revoke execute on function public.app_session_user(uuid), public.app_require(uuid, text[]), public.app_bump(text, text),
    public.records_fill_keys(), public.records_history() from public, anon, authenticated;

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
