-- MB System: run this once in Supabase → SQL Editor → New query → Run

create table if not exists public.documents (
  id          text primary key,
  owner       uuid not null default auth.uid() references auth.users(id) on delete cascade,
  type        text not null default 'quote',      -- quote | invoice | purchase
  client      text not null default '',
  doc_no      text,
  doc_date    date,
  total       numeric(14,2) not null default 0,
  data        jsonb not null,                      -- the full document as the app saves it
  updated_at  timestamptz not null default now()
);

create index if not exists documents_owner_idx  on public.documents(owner);
create index if not exists documents_client_idx on public.documents(owner, lower(client));
create index if not exists documents_date_idx   on public.documents(owner, doc_date);

-- Only the signed-in owner can see or change their documents
alter table public.documents enable row level security;

drop policy if exists "owner reads"   on public.documents;
drop policy if exists "owner inserts" on public.documents;
drop policy if exists "owner updates" on public.documents;
drop policy if exists "owner deletes" on public.documents;

create policy "owner reads"   on public.documents for select using (owner = auth.uid());
create policy "owner inserts" on public.documents for insert with check (owner = auth.uid());
create policy "owner updates" on public.documents for update using (owner = auth.uid()) with check (owner = auth.uid());
create policy "owner deletes" on public.documents for delete using (owner = auth.uid());

-- Live updates between devices
alter publication supabase_realtime add table public.documents;

-- Handy reports you can open in Supabase → Table editor / SQL (optional)
create or replace view public.monthly_profit with (security_invoker = true) as
with s as (
  select id, client, doc_date, coalesce((data->>'sub')::numeric, total) as sale
  from public.documents where type = 'invoice'
), p as (
  select data->>'linkTo' as link, coalesce((data->>'sub')::numeric, total) as cost
  from public.documents where type = 'purchase' and coalesce(data->>'linkTo','') <> ''
)
select to_char(s.doc_date, 'YYYY-MM') as month,
       sum(s.sale) as sales,
       sum(coalesce(c.cost,0)) as purchases,
       sum(s.sale - coalesce(c.cost,0)) as profit
from s join (select link, sum(cost) cost from p group by link) c on c.link = s.id
group by 1 order by 1 desc;
