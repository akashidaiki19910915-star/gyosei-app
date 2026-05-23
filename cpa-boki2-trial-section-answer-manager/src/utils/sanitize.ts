export function safeText(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value);
}
