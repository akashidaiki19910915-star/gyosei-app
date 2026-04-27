create table if not exists case_documents (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  case_id uuid,
  document_name text not null,
  status text default '未回収',
  received_date date,
  checked_date date,
  memo text,
  file_url text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

alter table case_documents enable row level security;

drop policy if exists "case_documents_own" on case_documents;

create policy "case_documents_own" on case_documents
for all
using (auth.uid() = user_id)
with check (auth.uid() = user_id);
