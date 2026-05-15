-- Add workflow fields for next action alerts (cases / daily_reports)
alter table if exists public.cases
  add column if not exists next_action_status text default 'open',
  add column if not exists next_action_completed_at timestamptz;

alter table if exists public.daily_reports
  add column if not exists next_action_status text default 'open',
  add column if not exists next_action_completed_at timestamptz;
