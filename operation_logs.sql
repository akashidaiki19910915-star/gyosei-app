create table if not exists operation_logs (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  action_type text not null,
  target_type text,
  target_id uuid,
  target_name text,
  detail text,
  created_at timestamptz default now()
);

alter table operation_logs enable row level security;

drop policy if exists "operation_logs_own" on operation_logs;

create policy "operation_logs_own" on operation_logs
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);
