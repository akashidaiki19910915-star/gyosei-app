import type { AnswerState, ExamSetRecord, HistoryEntry, MasteryMapItem, MistakeCard, ProblemDefinition, RecoveryPlan, StudyQueueItem } from '../types';
import { formatAmount, isAmountColumn } from './numberFormat';
import { draftSummary } from './scoring';

function escapeCsv(value: unknown): string {
  const raw = value === null || value === undefined ? '' : value;
  const text = String(raw).replace(/"/g, '""');
  return `"${text}"`;
}

function formatAnswerCell(columnName: string, value: string): string {
  return isAmountColumn(columnName) ? formatAmount(value) : value;
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
  answer.rows.forEach((row, index) => rows.push([index + 1, ...answer.columns.map((column, colIndex) => formatAnswerCell(column, row.cells[colIndex] ?? '')), row.grade, row.points]));
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
  const counts = new Map<string, { count: number; problemIds: Set<string>; latestAt: string }>();
  histories.forEach((history) => history.missReasons.forEach((reason) => {
    const current = counts.get(reason) ?? { count: 0, problemIds: new Set<string>(), latestAt: '' };
    current.count += 1;
    current.problemIds.add(history.displayId);
    if (history.savedAt > current.latestAt) current.latestAt = history.savedAt;
    counts.set(reason, current);
  }));
  return [['ミス原因', '件数', '該当問題ID', '最新発生日'], ...Array.from(counts.entries()).map(([reason, info]) => [reason, info.count, Array.from(info.problemIds).join('/'), info.latestAt])];
}

export function studyQueueCsvRows(items: StudyQueueItem[]): unknown[][] {
  return [
    ['優先順位', '状態ラベル', '科目', '大問対策', '問題ID', '論点名', '前回得点', '満点', '前回判定', '次回復習日', 'ミス原因', '前回メモ要約', '理由'],
    ...items.map((item) => [item.priority, item.statusLabels.join('/'), item.problem.subject, item.problem.sectionLabel, item.problem.displayId, item.problem.topic, item.previousScore ?? '', item.maxScore ?? '', item.rank, item.nextReviewDate, item.missReasons.join('/'), item.memoSummary, item.reason]),
  ];
}

export function dashboardCsvRows(histories: HistoryEntry[]): unknown[][] {
  const byProblem = new Map<string, HistoryEntry[]>();
  histories.forEach((history) => byProblem.set(history.displayId, [...(byProblem.get(history.displayId) ?? []), history]));
  return [
    ['問題ID', '演習回数', '最新得点', '最高得点', '最低得点', '最新判定', '最新復習日', '最新ミス原因'],
    ...Array.from(byProblem.entries()).map(([displayId, rows]) => {
      const sorted = rows.sort((a, b) => a.savedAt.localeCompare(b.savedAt));
      const latest = sorted.at(-1)!;
      return [displayId, rows.length, latest.score, Math.max(...rows.map((row) => row.score)), Math.min(...rows.map((row) => row.score)), latest.rank, latest.nextReviewDate, latest.missReasons.join('/')];
    }),
  ];
}

export function examSetCsvRows(examSets: ExamSetRecord[]): unknown[][] {
  return [
    ['作成日時', 'セット名', '開始日時', '終了日時', '所要秒数', '選択問題ID', '合計点', '70点判定', 'メモ'],
    ...examSets.map((set) => [set.createdAt, set.name, set.startedAt, set.finishedAt, set.durationSeconds, set.selectedProblemIds.join('/'), set.totalScore, set.passLineReached ? '合格ライン到達' : '復習優先', set.memo]),
  ];
}

export function masteryMapCsvRows(items: MasteryMapItem[]): unknown[][] {
  return [
    ['問題ID', '科目', '大問対策', '論点名', '最新得点', '満点', '得点率', '最新判定', '演習回数', 'A回数', 'B回数', 'C回数', '最終演習日', '次回復習日', '到達状態', '次の推奨アクション'],
    ...items.map((item) => [item.problem.displayId, item.problem.subject, item.problem.sectionLabel, item.problem.topic, item.latestScore ?? '', item.maxScore ?? '', item.scoreRate ?? '', item.latestRank, item.attempts, item.aCount, item.bCount, item.cCount, item.lastPracticedAt, item.nextReviewDate, item.status, item.action]),
  ];
}

export function mistakeCardCsvRows(cards: MistakeCard[], problemResolver: (problemId: string) => ProblemDefinition): unknown[][] {
  return [
    ['作成日', '更新日', '問題ID', '論点名', '重要度', '正しい処理の流れ', 'なぜ間違えたか', '次回最初に見る注意点', '次回確認事項', '同じミスを防ぐ一言', '関連ミス原因'],
    ...cards.map((card) => {
      const problem = problemResolver(card.problemId);
      return [card.createdAt, card.updatedAt, problem.displayId, problem.topic, card.importance, card.correctFlow, card.mistakeReason, card.firstReviewPoint, card.preSolveChecklist, card.preventionPhrase, card.missReasons.join('/')];
    }),
  ];
}

export function recoveryPlanCsvRows(plan: RecoveryPlan): unknown[][] {
  return [
    ['判定元', '合計点', '満点', '70点まで', '72点まで', '80点まで', '最大失点大問', '最多ミス原因', '次回重点対策', '取り戻し候補'],
    [plan.source, plan.totalScore, plan.maxScore, plan.shortage70, plan.shortage72, plan.shortage80, plan.biggestLossSection, plan.mostFrequentMissReason, plan.nextFocus, plan.recoveryCandidates.join('/')],
    [],
    ['大問', '得点', '満点', '失点'],
    ...plan.sectionRows.map((row) => [row.sectionLabel, row.score, row.maxScore, row.lostPoints]),
  ];
}
