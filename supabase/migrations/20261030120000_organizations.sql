-- Organizations: groups of users (a company). Each user belongs to at most one. Created by super admins (see
-- *_super_admins.sql), who assign users; organization admins add and remove their own members. Members are also
-- joined automatically:
--   - an invitation (organization_invites) for the user's email is accepted on their next visit, and
--   - a user who signs in with Microsoft whose email domain is the organization's `domain` joins it.
-- A mailbox added for the first time takes its settings from the same mailbox of another member of the same
-- organization (see mailboxes/index.ts). Not to be confused with ms_connections (a tenant's API keys).
-- Written only by the `organizations` edge function (service role); members may read their own organization.
create table if not exists public.organizations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  -- Email domain whose Microsoft sign-ins join automatically (lowercase, e.g. contoso.com); null: invitations only.
  domain text unique,
  created_by uuid references auth.users (id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.organization_members (
  user_id uuid primary key references auth.users (id) on delete cascade,
  org_id uuid not null references public.organizations (id) on delete cascade,
  role text not null default 'member' check (role in ('admin', 'member')),
  created_at timestamptz not null default now()
);
create index if not exists organization_members_org_idx on public.organization_members (org_id);

-- People invited before they have an account (or before their next visit); email is lowercase.
create table if not exists public.organization_invites (
  org_id uuid not null references public.organizations (id) on delete cascade,
  email text not null,
  role text not null default 'member' check (role in ('admin', 'member')),
  invited_by uuid references auth.users (id) on delete set null,
  created_at timestamptz not null default now(),
  primary key (org_id, email)
);
create index if not exists organization_invites_email_idx on public.organization_invites (email);

-- Members who were removed or left: not joined again by domain (an admin adding them back clears this).
create table if not exists public.organization_removed (
  org_id uuid not null references public.organizations (id) on delete cascade,
  user_id uuid not null references auth.users (id) on delete cascade,
  created_at timestamptz not null default now(),
  primary key (org_id, user_id)
);

-- The caller's organization, for RLS without recursing into organization_members' own policy.
create or replace function public.my_org_id()
returns uuid
language sql stable security definer set search_path = '' as $$
  select org_id from public.organization_members where user_id = (select auth.uid())
$$;

revoke all on function public.my_org_id() from public, anon;
grant execute on function public.my_org_id() to authenticated;

alter table public.organizations enable row level security;
alter table public.organization_members enable row level security;
alter table public.organization_invites enable row level security;
-- No policies: service role only.
alter table public.organization_removed enable row level security;
revoke all on public.organization_removed from anon, authenticated;

drop policy if exists "Members can read their organization" on public.organizations;
drop policy if exists "Members can read their organization's members" on public.organization_members;
drop policy if exists "Members can read their organization's invites" on public.organization_invites;
create policy "Members can read their organization"
  on public.organizations for select to authenticated using (id = (select public.my_org_id()));
create policy "Members can read their organization's members"
  on public.organization_members for select to authenticated using (org_id = (select public.my_org_id()));
create policy "Members can read their organization's invites"
  on public.organization_invites for select to authenticated using (org_id = (select public.my_org_id()));

revoke all on public.organizations, public.organization_members, public.organization_invites from anon, authenticated;
grant select on public.organizations, public.organization_members, public.organization_invites to authenticated;
