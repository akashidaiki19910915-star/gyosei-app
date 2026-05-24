import type { GradingStatus, HistoryEntry, ReviewRank, ReviewState } from '../types';
import { addDaysIso, nowIso } from './dates';

function intervalFor(status: GradingStatus, correctStreak: number, rank: ReviewRank, examSoon: boolean): number {
  let days = 7;
  if (status === '不正解' || rank === 'C') days = 1;
  else if (status === '部分正解') days = 2;
  else if (status === '迷いあり' || rank === 'B') days = 3;
  else if (status === '正解') {
    if (correctStreak >= 4) days = 60;
    else if (correctStreak >= 3) days = 30;
    else if (correctStreak >= 2) days = 14;
    else days = 7;
  }
  if (examSoon) return Math.min(days, 7);
  return days;
}

export function rankFromSelfGrading(status: GradingStatus, score: number, maxScore: number): ReviewRank {
  const rate = maxScore > 0 ? score / maxScore : 0;
  if (status === '不正解' || rate < 0.5) return 'C';
  if (status === '部分正解' || status === '迷いあり' || rate < 0.8) return 'B';
  return 'A';
}

export function updateReviewStateFromGrading(params: {
  problemId: string;
  previous?: ReviewState;
  status: GradingStatus;
  score: number;
  maxScore: number;
  examSoon?: boolean;
}): ReviewState {
  const rank = rankFromSelfGrading(params.status, params.score, params.maxScore);
  const correct = params.status === '正解';
  const correctStreak = correct ? (params.previous?.correctStreak ?? 0) + 1 : 0;
  const wrongStreak = correct ? 0 : (params.previous?.wrongStreak ?? 0) + 1;
  const intervalDays = intervalFor(params.status, correctStreak, rank, Boolean(params.examSoon));
  const now = nowIso();
  return {
    id: params.problemId,
    problemId: params.problemId,
    correctStreak,
    wrongStreak,
    lastResult: params.status,
    lastReviewedAt: now,
    nextReviewDate: addDaysIso(intervalDays),
    intervalDays,
    easeLevel: Math.max(1, Math.min(5, (params.previous?.easeLevel ?? 3) + (correct ? 1 : -1))),
    updatedAt: now,
  };
}

export function latestHistoryByProblem(histories: HistoryEntry[]): Map<string, HistoryEntry> {
  const map = new Map<string, HistoryEntry>();
  histories.forEach((history) => {
    const current = map.get(history.problemId);
    if (!current || history.savedAt > current.savedAt) map.set(history.problemId, history);
  });
  return map;
}
