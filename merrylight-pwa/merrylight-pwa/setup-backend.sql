-- Merrylight ERP: run this once in Supabase (SQL Editor -> New query -> paste -> Run)
create table if not exists kv (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);
alter table kv enable row level security;
create policy "mlerp read"   on kv for select to anon using (true);
create policy "mlerp insert" on kv for insert to anon with check (true);
create policy "mlerp update" on kv for update to anon using (true) with check (true);
create policy "mlerp delete" on kv for delete to anon using (true);
