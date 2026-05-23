import type { AnswerState, HistoryEntry, ProblemDefinition } from '../types';
import { draftSummary } from './scoring';

function escapeCsv(value: unknown): string {
  const raw = value === null || value === undefined ? '' : value;
  const text = String(raw).replace(/"/g, '""');
  return `"${text}"`;
}

export function downloadText(filename: string, text: string, type = 'text/plain;charset=utf-8'): void {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function downloadCsv(filename: string, rows: unknown[][]): void {
  const csv = rows.map((row) => row.map(escapeCsv).join(',')).join('\n');
  downloadText(filename, `\uFEFF${csv}`, 'text/csv;charset=utf-8');
}

export function currentAnswerCsvRows(answer: AnswerState, problem: ProblemDefinition): unknown[][] {
  const rows: unknown[][] = [
    ['保存日時', '科目', '大問対策', '問題ID', '論点名', '使用テンプレート', '問題得点', '満点', '行別得点合計', '採点済み', '採点日時', 'A/B/C判定', '次回復習日', 'ミス原因', '復習メモ', '下書きメモ要約'],
    [answer.updatedAt, problem.subject, problem.sectionLabel, problem.displayId, problem.topic, answer.templateName, answer.score, answer.maxScore, answer.rowPointsTotal, answer.scored ? '採点済み' : '未確定', answer.scoredAt, answer.rank, answer.nextReviewDate, answer.missReasons.join('/'), answer.reviewMemo, draftSummary(answer.draftMemo)],
    [],
    ['行番号', ...answer.columns, '採点', '行別得点'],
  ];
  answer.rows.forEach((row, index) => rows.push([index + 1, ...row.cells, row.grade, row.points]));
  return rows;
}

export function historyCsvRows(histories: HistoryEntry[]): unknown[][] {
  const rows: unknown[][] = [[
    '保存日時', '科目', '大問対策', '問題ID', '論点名', '使用テンプレート', '問題得点', '満点', '行別得点合計', '採点済み', '採点日時', 'A/B/C判定', '次回復習日', 'ミス原因', '復習メモ', '下書きメモ要約',
  ]];
  histories.forEach((history) => rows.push([
    history.savedAt, history.subject, history.sectionLabel, history.displayId, history.topic, history.templateName, history.score, history.maxScore, history.rowPointsTotal, history.scored ? '採点済み' : '未確定', history.scoredAt, history.rank, history.nextReviewDate, history.missReasons.join('/'), history.reviewMemo, history.draftMemoSummary,
  ]));
  return rows;
}

export function reviewTargetCsvRows(histories: HistoryEntry[]): unknown[][] {
  return historyCsvRows(histories);
}

export function problemStatsCsvRows(histories: HistoryEntry[]): unknown[][] {
  const latest = new Map<string, HistoryEntry>();
  histories.forEach((history) => latest.set(history.problemId, history));
  return historyCsvRows(Array.from(latest.values()));
}

export function missReasonCsvRows(histories: HistoryEntry[]): unknown[][] {
  const counts = new Map<string, number>();
  histories.forEach((history) => history.missReasons.forEach((reason) => counts.set(reason, (counts.get(reason) || 0) + 1)));
  return [['ミス原因', '件数'], ...Array.from(counts.entries())];
}
