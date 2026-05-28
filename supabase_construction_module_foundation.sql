-- 建設業許可モジュール 第1-A: DB・状態管理・保存基盤
-- Supabase SQL Editorで実行してください。

create table if not exists construction_case_details (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  case_id uuid not null references cases(id) on delete cascade,
  procedure_type text,
  permit_expiry_date date,
  fiscal_month integer,
  application_route text,
  memo text,
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  constraint construction_case_details_user_case_unique unique (user_id, case_id),
  constraint construction_case_details_fiscal_month_check check (fiscal_month is null or (fiscal_month between 1 and 12))
);

create table if not exists business_resource_templates (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null,
  procedure_type text,
  resource_category text,
  resource_name text not null,
  description text,
  url text,
  file_url text,
  customer_visible boolean not null default false,
  required boolean not null default false,
  last_checked_at date,
  source_type text,
  source_name text,
  memo text,
  related_task_type text,
  related_estimate_item text,
  related_expense_item text,
  sort_order integer default 0,
  is_active boolean not null default true,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create or replace function construction_module_set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists set_construction_case_details_updated_at on construction_case_details;
create trigger set_construction_case_details_updated_at
before update on construction_case_details
for each row execute function construction_module_set_updated_at();

drop trigger if exists set_business_resource_templates_updated_at on business_resource_templates;
create trigger set_business_resource_templates_updated_at
before update on business_resource_templates
for each row execute function construction_module_set_updated_at();

alter table construction_case_details enable row level security;
alter table business_resource_templates enable row level security;

drop policy if exists "construction_case_details_select_own" on construction_case_details;
drop policy if exists "construction_case_details_insert_own" on construction_case_details;
drop policy if exists "construction_case_details_update_own" on construction_case_details;
drop policy if exists "construction_case_details_delete_own" on construction_case_details;

create policy "construction_case_details_select_own" on construction_case_details
for select
using (auth.uid() = user_id);

create policy "construction_case_details_insert_own" on construction_case_details
for insert
with check (auth.uid() = user_id);

create policy "construction_case_details_update_own" on construction_case_details
for update
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

create policy "construction_case_details_delete_own" on construction_case_details
for delete
using (auth.uid() = user_id);

drop policy if exists "business_resource_templates_select_own" on business_resource_templates;
drop policy if exists "business_resource_templates_insert_own" on business_resource_templates;
drop policy if exists "business_resource_templates_update_own" on business_resource_templates;
drop policy if exists "business_resource_templates_delete_own" on business_resource_templates;

create policy "business_resource_templates_select_own" on business_resource_templates
for select
using (auth.uid() = user_id);

create policy "business_resource_templates_insert_own" on business_resource_templates
for insert
with check (auth.uid() = user_id);

create policy "business_resource_templates_update_own" on business_resource_templates
for update
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

create policy "business_resource_templates_delete_own" on business_resource_templates
for delete
using (auth.uid() = user_id);

create index if not exists idx_construction_case_details_user_id on construction_case_details(user_id);
create index if not exists idx_construction_case_details_case_id on construction_case_details(case_id);
create index if not exists idx_construction_case_details_procedure_type on construction_case_details(procedure_type);
create index if not exists idx_business_resource_templates_user_id on business_resource_templates(user_id);
create index if not exists idx_business_resource_templates_procedure_type on business_resource_templates(procedure_type);
create index if not exists idx_business_resource_templates_active_order on business_resource_templates(user_id, is_active, sort_order);
