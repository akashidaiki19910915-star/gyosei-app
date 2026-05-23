import { problemCatalog } from '../data/problemCatalog';
import type { HistoryEntry, ProblemDefinition, StudyQueueItem } from '../types';
import { isDueTodayOrEarlier, isOverdue } from './dates';

function latestByProblem(histories: HistoryEntry[]): Map<string, HistoryEntry> {
  const map = new Map<string, HistoryEntry>();
  histories.forEach((history) => {
    const current = map.get(history.problemId);
    if (!current || history.savedAt > current.savedAt) map.set(history.problemId, history);
  });
  return map;
}

function buildItem(problem: ProblemDefinition, history: HistoryEntry | undefined, priority: number, reason: string, statusLabels: StudyQueueItem['statusLabels']): StudyQueueItem {
  return {
    priority,
    statusLabels,
    problem,
    latestHistory: history,
    previousScore: history?.score,
    maxScore: history?.maxScore,
    rank: history?.rank ?? '',
    nextReviewDate: history?.nextReviewDate ?? '',
    missReasons: history?.missReasons ?? [],
    memoSummary: history?.draftMemoSummary || history?.reviewMemo?.slice(0, 80) || '',
    reason,
  };
}

export function buildStudyQueue(histories: HistoryEntry[]): StudyQueueItem[] {
  const latest = latestByProblem(histories);
  const items: StudyQueueItem[] = [];
  const pushed = new Set<string>();

  function add(problem: ProblemDefinition, history: HistoryEntry | undefined, priority: number, reason: string, labels: StudyQueueItem['statusLabels']) {
    if (pushed.has(problem.id)) return;
    pushed.add(problem.id);
    items.push(buildItem(problem, history, priority, reason, labels));
  }

  problemCatalog.forEach((problem) => {
    const history = latest.get(problem.id);
    if (!history) return;
    const overdue = isOverdue(history.nextReviewDate);
    const due = isDueTodayOrEarlier(history.nextReviewDate);
    const noDateC = history.rank === 'C' && !history.nextReviewDate;
    if (history.rank === 'C' && overdue) add(problem, history, 1, '期限超過のC判定', ['期限超過', 'C判定']);
    else if (history.rank === 'C' && due) add(problem, history, 2, '今日が復習日のC判定', ['今日復習', 'C判定']);
    else if (history.rank === 'B' && overdue) add(problem, history, 3, '期限超過のB判定', ['期限超過', 'B判定']);
    else if (history.rank === 'B' && due) add(problem, history, 4, '今日が復習日のB判定', ['今日復習', 'B判定']);
    else if (history.rank === 'A' && overdue) add(problem, history, 5, '期限超過のA判定', ['期限超過', 'A判定']);
    else if (history.rank === 'A' && due) add(problem, history, 6, '今日が復習日のA判定', ['今日復習', 'A判定']);
    else if (noDateC) add(problem, history, 6.5, 'C判定だが次回復習日が未設定', ['日付未設定C', 'C判定']);
  });

  problemCatalog.forEach((problem) => {
    if (!latest.has(problem.id)) add(problem, undefined, 7, '未着手問題', ['未着手']);
  });

  problemCatalog.forEach((problem) => {
    const history = latest.get(problem.id);
    if (!history || pushed.has(problem.id)) return;
    const rate = history.maxScore ? history.score / history.maxScore : 1;
    if (rate < 0.6) add(problem, history, 8, '直近の得点率が低い問題', ['低得点']);
  });

  problemCatalog.forEach((problem) => {
    const history = latest.get(problem.id);
    if (!history || pushed.has(problem.id)) return;
    if ((history.missReasons?.length ?? 0) >= 2) add(problem, history, 9, 'ミス原因が多い問題', ['ミス多い']);
  });

  return items.sort((a, b) => a.priority - b.priority || a.problem.displayId.localeCompare(b.problem.displayId));
}

export function queueCounts(items: StudyQueueItem[]) {
  return {
    total: items.length,
    overdue: items.filter((item) => item.statusLabels.includes('期限超過')).length,
    c: items.filter((item) => item.statusLabels.includes('C判定')).length,
    untouched: items.filter((item) => item.statusLabels.includes('未着手')).length,
  };
}
