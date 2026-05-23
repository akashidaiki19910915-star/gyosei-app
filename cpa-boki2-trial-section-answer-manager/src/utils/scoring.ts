import type { AnswerRow, AnswerState } from '../types';

export function sumRowPoints(rows: AnswerRow[]): number {
  return rows.reduce((sum, row) => sum + (Number(row.points) || 0), 0);
}

export function hasMeaningfulAnswer(answer: AnswerState): boolean {
  const hasCells = answer.rows.some((row) => row.cells.some((cell) => cell.trim() !== ''));
  const hasScoring = answer.rows.some((row) => row.grade !== '未採点' || Number(row.points) !== 0);
  return hasCells || hasScoring || answer.score !== 0 || answer.rank !== '' || answer.missReasons.length > 0 || answer.draftMemo.trim() !== '' || answer.reviewMemo.trim() !== '' || answer.scored;
}

export function draftSummary(value: string, length = 80): string {
  return value.replaceAll('\n', ' ').replaceAll('\t', ' ').trim().slice(0, length);
}
