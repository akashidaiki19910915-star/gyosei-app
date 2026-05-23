import { problemCatalog } from '../data/problemCatalog';
import type { DashboardProblemStats, HistoryEntry, MissReason, SectionId } from '../types';
import { isDueTodayOrEarlier, isOverdue } from './dates';

function scoreRate(history: HistoryEntry): number {
  return history.maxScore ? history.score / history.maxScore : 0;
}

export function buildProblemStats(histories: HistoryEntry[]): DashboardProblemStats[] {
  return problemCatalog.map((problem) => {
    const rows = histories.filter((history) => history.problemId === problem.id).sort((a, b) => a.savedAt.localeCompare(b.savedAt));
    const latest = rows.at(-1);
    const previous = rows.at(-2);
    let trend: DashboardProblemStats['trend'] = 'データ不足';
    if (latest && previous) {
      if (latest.score > previous.score) trend = '改善';
      else if (latest.score < previous.score) trend = '悪化';
      else trend = '横ばい';
    }
    return {
      problem,
      attempts: rows.length,
      latestScore: latest?.score ?? null,
      highestScore: rows.length ? Math.max(...rows.map((row) => row.score)) : null,
      lowestScore: rows.length ? Math.min(...rows.map((row) => row.score)) : null,
      latestRank: latest?.rank ?? '',
      latestMissReasons: latest?.missReasons ?? [],
      latestReviewDate: latest?.nextReviewDate ?? '',
      trend,
    };
  });
}

export function buildDashboard(histories: HistoryEntry[], answerCount: number) {
  const rates = histories.map(scoreRate).filter((rate) => Number.isFinite(rate));
  const averageRate = rates.length ? rates.reduce((sum, value) => sum + value, 0) / rates.length : 0;
  const averageScore = histories.length ? histories.reduce((sum, history) => sum + history.score, 0) / histories.length : 0;
  const problemStats = buildProblemStats(histories);
  const untouchedCount = problemStats.filter((row) => row.attempts === 0).length;
  const overdueCount = histories.filter((history) => isOverdue(history.nextReviewDate)).length;
  const todayCount = histories.filter((history) => isDueTodayOrEarlier(history.nextReviewDate)).length;
  const missCounts = new Map<MissReason, { count: number; problemIds: Set<string>; latestAt: string }>();

  histories.forEach((history) => {
    history.missReasons.forEach((reason) => {
      const current = missCounts.get(reason) ?? { count: 0, problemIds: new Set<string>(), latestAt: '' };
      current.count += 1;
      current.problemIds.add(history.displayId);
      if (history.savedAt > current.latestAt) current.latestAt = history.savedAt;
      missCounts.set(reason, current);
    });
  });

  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

  const inRange = (history: HistoryEntry, date: Date) => new Date(history.savedAt) >= date;

  const sectionStats = (['q1', 'q2', 'q3', 'q4', 'q5'] as SectionId[]).map((sectionId) => {
    const sectionRows = histories.filter((history) => history.sectionId === sectionId);
    const count = sectionRows.length;
    const cCount = sectionRows.filter((history) => history.rank === 'C').length;
    const reasons = new Map<string, number>();
    sectionRows.forEach((history) => history.missReasons.forEach((reason) => reasons.set(reason, (reasons.get(reason) || 0) + 1)));
    const commonReason = Array.from(reasons.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] ?? '-';
    return {
      sectionId,
      sectionLabel: problemCatalog.find((problem) => problem.sectionId === sectionId)?.sectionLabel ?? sectionId,
      attempts: count,
      averageScore: count ? sectionRows.reduce((sum, history) => sum + history.score, 0) / count : 0,
      cCount,
      commonReason,
    };
  });

  return {
    summary: {
      totalAttempts: histories.length,
      historyCount: histories.length,
      answerCount,
      averageScore,
      averageRate,
      passLikeCount: histories.filter((history) => history.maxScore && history.score / history.maxScore >= 0.7).length,
      aCount: histories.filter((history) => history.rank === 'A').length,
      bCount: histories.filter((history) => history.rank === 'B').length,
      cCount: histories.filter((history) => history.rank === 'C').length,
      untouchedCount,
      overdueCount,
      todayCount,
      sevenDayAttempts: histories.filter((history) => inRange(history, sevenDaysAgo)).length,
      thirtyDayAttempts: histories.filter((history) => inRange(history, thirtyDaysAgo)).length,
      sevenDayC: histories.filter((history) => history.rank === 'C' && inRange(history, sevenDaysAgo)).length,
      thirtyDayC: histories.filter((history) => history.rank === 'C' && inRange(history, thirtyDaysAgo)).length,
    },
    problemStats,
    sectionStats,
    missReasonRanking: Array.from(missCounts.entries()).map(([reason, info]) => ({ reason, count: info.count, problemIds: Array.from(info.problemIds), latestAt: info.latestAt })).sort((a, b) => b.count - a.count),
  };
}
