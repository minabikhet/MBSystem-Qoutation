-- Optional: keeps the monthly_profit report in Supabase in line with the app
-- (ignores deleted documents). Run once in Supabase → SQL Editor.
create or replace view public.monthly_profit with (security_invoker = true) as
with d as (
  select * from public.documents where coalesce((data->>'deleted')::boolean, false) = false
), s as (
  select id, client, doc_date, coalesce((data->>'sub')::numeric, total) as sale from d where type = 'invoice'
), p as (
  select data->>'linkTo' as link, coalesce((data->>'sub')::numeric, total) as cost
  from d where type = 'purchase' and coalesce(data->>'linkTo','') <> ''
)
select to_char(s.doc_date, 'YYYY-MM') as month,
       sum(s.sale) as sales, sum(c.cost) as purchases, sum(s.sale - c.cost) as profit
from s join (select link, sum(cost) cost from p group by link) c on c.link = s.id
group by 1 order by 1 desc;
