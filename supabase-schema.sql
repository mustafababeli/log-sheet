create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  email text unique,
  full_name text,
  created_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.facilities (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  created_by uuid not null references public.profiles (id) on delete cascade,
  created_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.facility_users (
  facility_id uuid not null references public.facilities (id) on delete cascade,
  user_id uuid not null references public.profiles (id) on delete cascade,
  role text not null default 'member' check (role in ('admin', 'member')),
  created_at timestamptz not null default timezone('utc', now()),
  primary key (facility_id, user_id)
);

create table if not exists public.operators (
  id uuid primary key default gen_random_uuid(),
  facility_id uuid not null references public.facilities (id) on delete cascade,
  name text not null,
  created_at timestamptz not null default timezone('utc', now())
);

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'operators_id_facility_id_key'
      and conrelid = 'public.operators'::regclass
  ) then
    alter table public.operators
      add constraint operators_id_facility_id_key unique (id, facility_id);
  end if;
end
$$;

create table if not exists public.daily_logs (
  id uuid primary key default gen_random_uuid(),
  facility_id uuid not null references public.facilities (id) on delete cascade,
  operator_id uuid not null references public.operators (id) on delete cascade,
  log_date date not null,
  entries jsonb not null default '{}'::jsonb,
  skipped_fields jsonb not null default '[]'::jsonb,
  created_by uuid not null references public.profiles (id) on delete cascade,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (facility_id, operator_id, log_date)
);

create table if not exists public.monthly_readings (
  id uuid primary key default gen_random_uuid(),
  facility_id uuid not null references public.facilities (id) on delete cascade,
  operator_id uuid not null references public.operators (id) on delete cascade,
  period text not null,
  entries jsonb not null default '{}'::jsonb,
  photos jsonb not null default '{}'::jsonb,
  created_by uuid not null references public.profiles (id) on delete cascade,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (facility_id, operator_id, period)
);

do $$
begin
  if exists (
    select 1
    from pg_constraint
    where conname = 'daily_logs_operator_id_fkey'
      and conrelid = 'public.daily_logs'::regclass
  ) then
    alter table public.daily_logs drop constraint daily_logs_operator_id_fkey;
  end if;

  if not exists (
    select 1
    from pg_constraint
    where conname = 'daily_logs_operator_facility_fkey'
      and conrelid = 'public.daily_logs'::regclass
  ) then
    alter table public.daily_logs
      add constraint daily_logs_operator_facility_fkey
      foreign key (operator_id, facility_id)
      references public.operators (id, facility_id)
      on delete cascade;
  end if;

  if exists (
    select 1
    from pg_constraint
    where conname = 'monthly_readings_operator_id_fkey'
      and conrelid = 'public.monthly_readings'::regclass
  ) then
    alter table public.monthly_readings drop constraint monthly_readings_operator_id_fkey;
  end if;

  if not exists (
    select 1
    from pg_constraint
    where conname = 'monthly_readings_operator_facility_fkey'
      and conrelid = 'public.monthly_readings'::regclass
  ) then
    alter table public.monthly_readings
      add constraint monthly_readings_operator_facility_fkey
      foreign key (operator_id, facility_id)
      references public.operators (id, facility_id)
      on delete cascade;
  end if;
end
$$;

insert into storage.buckets (id, name, public)
values ('monthly-photos', 'monthly-photos', false)
on conflict (id) do update
set public = excluded.public;

alter table public.profiles enable row level security;
alter table public.facilities enable row level security;
alter table public.facility_users enable row level security;
alter table public.operators enable row level security;
alter table public.daily_logs enable row level security;
alter table public.monthly_readings enable row level security;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email)
  values (new.id, new.email)
  on conflict (id) do update
    set email = excluded.email;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();

create or replace function public.is_facility_admin(target_facility_id uuid)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.facility_users
    where facility_id = target_facility_id
      and user_id = auth.uid()
      and role = 'admin'
  );
$$;

create or replace function public.create_facility_with_membership(facility_name text)
returns uuid
language plpgsql
security definer
set search_path = public
as $$
declare
  new_facility_id uuid;
  current_user_id uuid;
begin
  current_user_id := auth.uid();

  if current_user_id is null then
    raise exception 'Authentication required';
  end if;

  if facility_name is null or btrim(facility_name) = '' then
    raise exception 'Facility name is required';
  end if;

  insert into public.facilities (name, created_by)
  values (btrim(facility_name), current_user_id)
  returning id into new_facility_id;

  insert into public.facility_users (facility_id, user_id, role)
  values (new_facility_id, current_user_id, 'admin');

  return new_facility_id;
end;
$$;

drop policy if exists "profiles are viewable by signed in users" on public.profiles;
create policy "profiles are viewable by signed in users"
on public.profiles
for select
to authenticated
using (auth.uid() = id);

create policy "users can update their own profile"
on public.profiles
for update
to authenticated
using (auth.uid() = id)
with check (auth.uid() = id);

create policy "users can view assigned facilities"
on public.facilities
for select
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = facilities.id
      and facility_users.user_id = auth.uid()
  )
);

drop policy if exists "users can create facilities" on public.facilities;

create policy "facility admins can update facilities"
on public.facilities
for update
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = facilities.id
      and facility_users.user_id = auth.uid()
      and facility_users.role = 'admin'
  )
)
with check (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = facilities.id
      and facility_users.user_id = auth.uid()
      and facility_users.role = 'admin'
  )
);

create policy "facility admins can delete facilities"
on public.facilities
for delete
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = facilities.id
      and facility_users.user_id = auth.uid()
      and facility_users.role = 'admin'
  )
);

create policy "users can view facility memberships"
on public.facility_users
for select
to authenticated
using (
  user_id = auth.uid()
  or public.is_facility_admin(facility_id)
);

drop policy if exists "facility admins can manage memberships" on public.facility_users;
drop policy if exists "users can create their own facility memberships" on public.facility_users;

create policy "facility admins can update memberships"
on public.facility_users
for update
to authenticated
using (public.is_facility_admin(facility_id))
with check (public.is_facility_admin(facility_id));

create policy "facility admins can delete memberships"
on public.facility_users
for delete
to authenticated
using (public.is_facility_admin(facility_id))
with check (public.is_facility_admin(facility_id));

create policy "users can view operators for assigned facilities"
on public.operators
for select
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = operators.facility_id
      and facility_users.user_id = auth.uid()
  )
);

create policy "facility admins can manage operators"
on public.operators
for all
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = operators.facility_id
      and facility_users.user_id = auth.uid()
      and facility_users.role = 'admin'
  )
)
with check (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = operators.facility_id
      and facility_users.user_id = auth.uid()
      and facility_users.role = 'admin'
  )
);

create policy "users can view daily logs for assigned facilities"
on public.daily_logs
for select
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = daily_logs.facility_id
      and facility_users.user_id = auth.uid()
  )
);

create policy "facility members can manage daily logs"
on public.daily_logs
for all
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = daily_logs.facility_id
      and facility_users.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = daily_logs.facility_id
      and facility_users.user_id = auth.uid()
  )
);

create policy "users can view monthly readings for assigned facilities"
on public.monthly_readings
for select
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = monthly_readings.facility_id
      and facility_users.user_id = auth.uid()
  )
);

create policy "facility members can manage monthly readings"
on public.monthly_readings
for all
to authenticated
using (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = monthly_readings.facility_id
      and facility_users.user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.facility_users
    where facility_users.facility_id = monthly_readings.facility_id
      and facility_users.user_id = auth.uid()
  )
);

drop policy if exists "users can view monthly photo objects" on storage.objects;
create policy "users can view monthly photo objects"
on storage.objects
for select
to authenticated
using (
  bucket_id = 'monthly-photos'
  and exists (
    select 1
    from public.facility_users
    where facility_users.facility_id::text = split_part(name, '/', 1)
      and facility_users.user_id = auth.uid()
  )
);

drop policy if exists "users can upload monthly photo objects" on storage.objects;
create policy "users can upload monthly photo objects"
on storage.objects
for insert
to authenticated
with check (
  bucket_id = 'monthly-photos'
  and exists (
    select 1
    from public.facility_users
    where facility_users.facility_id::text = split_part(name, '/', 1)
      and facility_users.user_id = auth.uid()
  )
);

drop policy if exists "users can update monthly photo objects" on storage.objects;
create policy "users can update monthly photo objects"
on storage.objects
for update
to authenticated
using (
  bucket_id = 'monthly-photos'
  and exists (
    select 1
    from public.facility_users
    where facility_users.facility_id::text = split_part(name, '/', 1)
      and facility_users.user_id = auth.uid()
  )
)
with check (
  bucket_id = 'monthly-photos'
  and exists (
    select 1
    from public.facility_users
    where facility_users.facility_id::text = split_part(name, '/', 1)
      and facility_users.user_id = auth.uid()
  )
);

drop policy if exists "users can delete monthly photo objects" on storage.objects;
create policy "users can delete monthly photo objects"
on storage.objects
for delete
to authenticated
using (
  bucket_id = 'monthly-photos'
  and exists (
    select 1
    from public.facility_users
    where facility_users.facility_id::text = split_part(name, '/', 1)
      and facility_users.user_id = auth.uid()
  )
);
