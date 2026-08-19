-- ============================================================
-- Nova Ops Dashboard — Supabase schema
-- Run this once in your Supabase project's SQL Editor
-- (Dashboard → SQL Editor → New query → paste → Run)
-- ============================================================

-- 1) Shared data store — mirrors the old localStorage keys 1:1, so all existing
--    tabs keep working with minimal rewiring. One row per data collection
--    (inventory, attendance, employees, finance, orders, etc.)
create table if not exists app_storage (
  key        text primary key,
  value      jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now()
);

alter table app_storage enable row level security;

-- Any authenticated (logged-in) user can read all shared data — tab-level
-- restriction happens in the UI layer, not by hiding rows here, since the
-- data isn't split per-tab at the database level (see note in README).
create policy "authenticated users can read app_storage"
  on app_storage for select
  to authenticated
  using (true);

create policy "authenticated users can write app_storage"
  on app_storage for all
  to authenticated
  using (true)
  with check (true);


-- 2) Per-user access control — who can log in, whether they're an admin
--    (sees everything), and which tabs they're allowed to see otherwise.
create table if not exists user_permissions (
  id           uuid primary key default gen_random_uuid(),
  email        text unique not null,
  display_name text not null default '',
  is_admin     boolean not null default false,
  allowed_tabs text[] not null default '{}',
  created_at   timestamptz not null default now()
);

alter table user_permissions enable row level security;

-- Everyone logged in can read the permissions table (needed so the app can
-- look up its own row after login, and so admins can see the full list).
create policy "authenticated users can read user_permissions"
  on user_permissions for select
  to authenticated
  using (true);

-- Only an existing admin can add/edit/remove permission rows.
create policy "admins can manage user_permissions"
  on user_permissions for all
  to authenticated
  using (
    exists (
      select 1 from user_permissions up
      where up.email = auth.jwt() ->> 'email' and up.is_admin = true
    )
  )
  with check (
    exists (
      select 1 from user_permissions up
      where up.email = auth.jwt() ->> 'email' and up.is_admin = true
    )
  );


-- 3) Seed yourself as the first admin — REPLACE the email below with the
--    address you'll log in with (create that user first in Authentication →
--    Users → Add User, then run this).
insert into user_permissions (email, display_name, is_admin, allowed_tabs)
values ('YOUR-EMAIL@example.com', 'Admin', true, '{}')
on conflict (email) do update set is_admin = true;
