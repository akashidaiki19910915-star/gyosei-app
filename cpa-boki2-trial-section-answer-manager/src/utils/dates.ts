import type { ReviewRank } from '../types';

export function todayIso(): string {
  return new Date().toISOString().slice(0, 10);
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function formatDateTime(value: string): string {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString('ja-JP');
}

export function addDaysIso(days: number): string {
  const date = new Date();
  date.setDate(date.getDate() + days);
  return date.toISOString().slice(0, 10);
}

export function reviewDateForRank(rank: ReviewRank): string {
  if (rank === 'A') return addDaysIso(7);
  if (rank === 'B') return addDaysIso(3);
  if (rank === 'C') return addDaysIso(1);
  return '';
}

function isDigits(value: string): boolean {
  return value.split('').every((char) => char >= '0' && char <= '9');
}

export function isValidIsoDate(value: string): boolean {
  if (!value || value.length !== 10) return false;
  const parts = value.split('-');
  if (parts.length !== 3) return false;
  return parts[0].length === 4 && parts[1].length === 2 && parts[2].length === 2 && parts.every(isDigits);
}

export function isDueTodayOrEarlier(value: string): boolean {
  if (!isValidIsoDate(value)) return false;
  return value <= todayIso();
}

export function isOverdue(value: string): boolean {
  if (!isValidIsoDate(value)) return false;
  return value < todayIso();
}
