alter table public.estimate_calculations
  add column if not exists keikan_reason text,
  add column if not exists sengi_reason text,
  add column if not exists zaisan_reason text;
