-- 第2-A: estimate_items.item_type 追加・見積明細区分対応
-- Supabase SQL EditorでPRマージ前（またはデプロイ前）に実行してください。

alter table estimate_items
add column if not exists item_type text;

alter table estimate_items
alter column item_type set default 'reward';

update estimate_items
set item_type = 'reward'
where item_type is null or item_type = '';

alter table estimate_items
alter column item_type set not null;

do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'estimate_items_item_type_check'
      and conrelid = 'estimate_items'::regclass
  ) then
    alter table estimate_items
    add constraint estimate_items_item_type_check
    check (item_type in ('reward', 'expense', 'advance', 'discount'));
  end if;
end $$;
