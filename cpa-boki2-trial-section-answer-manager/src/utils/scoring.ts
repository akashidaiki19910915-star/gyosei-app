import type { AnswerRow, AnswerState } from '../types';

export function sumRowPoints(rows: AnswerRow[]): number {
  return rows.reduce((sum, row) => sum + (Number(row.points) || 0), 0);
}

function rowsHaveMeaning(rows: AnswerRow[]): boolean {
  const hasCells = rows.some((row) => row.cells.some((cell) => cell.trim() !== ''));
  const hasScoring = rows.some((row) => row.grade !== '未採点' || Number(row.points) !== 0);
  return hasCells || hasScoring;
}

export function hasMeaningfulAnswer(answer: AnswerState): boolean {
  const hasBlockAnswer = answer.problemBlocks?.some((block) => rowsHaveMeaning(block.rows) || block.score !== 0 || block.rank !== '' || block.missReasons.length > 0 || block.memo.trim() !== '' || block.nextReviewPoint.trim() !== '') ?? false;
  return rowsHaveMeaning(answer.rows) || hasBlockAnswer || answer.score !== 0 || answer.rank !== '' || answer.missReasons.length > 0 || answer.draftMemo.trim() !== '' || answer.reviewMemo.trim() !== '' || answer.scored;
}

export function draftSummary(value: string, length = 80): string {
  return value.replaceAll('\n', ' ').replaceAll('\t', ' ').trim().slice(0, length);
}
