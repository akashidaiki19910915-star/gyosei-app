import { problemCatalog } from '../data/problemCatalog';
import type { HistoryEntry, MasteryFilter, MasteryMapItem } from '../types';
import { isOverdue } from './dates';

function byProblem(histories: HistoryEntry[]): Map<string, HistoryEntry[]> {
  const map = new Map<string, HistoryEntry[]>();
  histories.forEach((history) => {
    const rows = map.get(history.problemId) ?? [];
    rows.push(history);
    map.set(history.problemId, rows);
  });
  map.forEach((rows) => rows.sort((a, b) => a.savedAt.localeCompare(b.savedAt)));
  return map;
}

function scoreRate(score: number, maxScore: number): number | null {
  if (!maxScore) return null;
  return Math.round((score / maxScore) * 1000) / 10;
}

function isMastered(rows: HistoryEntry[]): boolean {
  if (rows.length < 2) return false;
  const latestTwo = rows.slice(-2);
  const twoA = latestTwo.every((row) => row.rank === 'A');
  const two80 = latestTwo.every((row) => scoreRate(row.score, row.maxScore) !== null && Number(scoreRate(row.score, row.maxScore)) >= 80);
  const aCount = rows.filter((row) => row.rank === 'A').length;
  const latest = rows.at(-1);
  const latestGood = latest?.rank === 'A' || latest?.rank === 'B';
  return twoA || two80 || (aCount >= 2 && Boolean(latestGood));
}

function actionFor(item: Omit<MasteryMapItem, 'action' | 'sortPriority'>): string {
  if (item.attempts === 0) return 'まず1回解く';
  if (isOverdue(item.nextReviewDate)) return '本日復習';
  if (item.latestRank === 'C') return '白紙再現・解き直しカードを見て再演習';
  if (item.scoreRate !== null && item.scoreRate < 70) return '失点原因を確認して再演習';
  if (item.latestRank === 'B') return 'もう1回解いてA化';
  if (item.status === '合格水準') return '試験前総復習まで保留';
  return '次の弱点問題へ進む';
}

function priorityFor(item: Omit<MasteryMapItem, 'action' | 'sortPriority'>): number {
  if (isOverdue(item.nextReviewDate) && item.latestRank === 'C') return 1;
  if (item.attempts === 0) return 2;
  if (item.latestRank === 'C') return 3;
  if (item.scoreRate !== null && item.scoreRate < 70) return 4;
  if (isOverdue(item.nextReviewDate) && item.latestRank === 'B') return 5;
  if (item.attempts <= 1) return 6;
  if (item.latestRank === 'B') return 7;
  if (item.status === '合格水準') return 8;
  return 9;
}

export function buildMasteryMap(histories: HistoryEntry[]): MasteryMapItem[] {
  const grouped = byProblem(histories);
  return problemCatalog.map((problem) => {
    const rows = grouped.get(problem.id) ?? [];
    const latest = rows.at(-1);
    const attempts = rows.length;
    const latestScore = latest?.score ?? null;
    const maxScore = latest?.maxScore ?? null;
    const rate = latest && maxScore !== null ? scoreRate(latest.score, latest.maxScore) : null;
    const aCount = rows.filter((row) => row.rank === 'A').length;
    const bCount = rows.filter((row) => row.rank === 'B').length;
    const cCount = rows.filter((row) => row.rank === 'C').length;
    const mastered = isMastered(rows);
    const overdue = latest ? isOverdue(latest.nextReviewDate) : false;
    const dangerous = attempts === 0 || latest?.rank === 'C' || (rate !== null && rate < 70) || overdue || cCount >= 2 || attempts <= 1;
    let status: MasteryMapItem['status'] = '未着手';
    if (overdue) status = '期限超過';
    else if (dangerous && attempts > 0) status = latest?.rank === 'C' || cCount >= 2 || (rate !== null && rate < 70) ? '危険問題' : '再復習対象';
    else if (mastered) status = '合格水準';
    else if (attempts >= 3) status = '3回以上演習済み';
    else if (attempts === 2) status = '2回演習済み';
    else if (attempts === 1) status = '1回演習済み';

    const base = {
      problem,
      latestScore,
      maxScore,
      scoreRate: rate,
      latestRank: latest?.rank ?? '',
      attempts,
      aCount,
      bCount,
      cCount,
      lastPracticedAt: latest?.savedAt ?? '',
      nextReviewDate: latest?.nextReviewDate ?? '',
      status,
    };
    const action = actionFor(base);
    const sortPriority = priorityFor(base);
    return { ...base, action, sortPriority };
  }).sort((a, b) => a.sortPriority - b.sortPriority || a.problem.displayId.localeCompare(b.problem.displayId, 'ja'));
}

export function filterMasteryMap(items: MasteryMapItem[], filter: MasteryFilter): MasteryMapItem[] {
  if (filter === '全体') return items;
  if (filter === '未着手') return items.filter((item) => item.status === '未着手');
  if (filter === '危険問題') return items.filter((item) => item.status === '危険問題' || item.status === '再復習対象' || item.status === '期限超過' || item.attempts === 0);
  if (filter === '期限超過') return items.filter((item) => item.status === '期限超過');
  if (filter === 'C判定') return items.filter((item) => item.latestRank === 'C');
  if (filter === 'B判定') return items.filter((item) => item.latestRank === 'B');
  if (filter === '合格水準') return items.filter((item) => item.status === '合格水準');
  if (filter === '商業簿記') return items.filter((item) => item.problem.subject === 'commercial');
  if (filter === '工業簿記') return items.filter((item) => item.problem.subject === 'industrial');
  return items.filter((item) => item.problem.sectionLabel === filter);
}
