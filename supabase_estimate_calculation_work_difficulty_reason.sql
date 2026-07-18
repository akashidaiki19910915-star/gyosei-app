alter table public.estimate_calculations
  add column if not exists document_level_reason text;
