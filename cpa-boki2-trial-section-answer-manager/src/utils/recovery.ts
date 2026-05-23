import { problemCatalog } from '../data/problemCatalog';
import type { ExamSetRecord, HistoryEntry, MissReason, RecoveryPlan, RecoveryPlanRow, SectionId } from '../types';

const sectionLabels: Record<SectionId, string> = {
  q1: '第1問対策',
  q2: '第2問対策',
  q3: '第3問対策',
  q4: '第4問対策',
  q5: '第5問対策',
};

function shortage(target: number, score: number): number {
  return Math.max(0, target - score);
}

function mostFrequentMissReason(histories: HistoryEntry[]): string {
  const counts = new Map<MissReason, number>();
  histories.forEach((history) => history.missReasons.forEach((reason) => counts.set(reason, (counts.get(reason) ?? 0) + 1)));
  const sorted = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]);
  return sorted[0]?.[0] ?? '未登録';
}

function focusFor(reason: string, cCount: number): string {
  if (cCount > 0) return 'C判定問題の白紙再現・解き直しカードを作ってから再演習';
  if (reason === '金額ミス') return '金額欄・集計欄を重点確認';
  if (reason === '仕訳ミス' || reason === '借方貸方逆') return '第1問対策と仕訳テーブルを優先';
  if (reason === '集計ミス') return '第3問または第5問の表問題を優先';
  if (reason === '時間不足') return '90分セット演習を優先';
  if (reason === '問題文読み落とし') return '設問番号・条件整理を最初に確認';
  if (reason === '表の入力位置ミス') return '解答欄形式と列位置を確認してから入力';
  return '直近の低得点問題から再演習';
}

function byLatestSectionHistories(histories: HistoryEntry[]): RecoveryPlanRow[] {
  const latestBySection = new Map<SectionId, HistoryEntry>();
  histories.forEach((history) => {
    const current = latestBySection.get(history.sectionId);
    if (!current || history.savedAt > current.savedAt) latestBySection.set(history.sectionId, history);
  });
  return (Object.keys(sectionLabels) as SectionId[]).map((sectionId) => {
    const latest = latestBySection.get(sectionId);
    return {
      sectionId,
      sectionLabel: sectionLabels[sectionId],
      score: latest?.score ?? 0,
      maxScore: latest?.maxScore ?? 20,
      lostPoints: Math.max(0, (latest?.maxScore ?? 20) - (latest?.score ?? 0)),
    };
  });
}

function byExamSet(record: ExamSetRecord): RecoveryPlanRow[] {
  return (Object.keys(sectionLabels) as SectionId[]).map((sectionId) => {
    const problem = problemCatalog.find((item) => item.sectionId === sectionId && record.selectedProblemIds.includes(item.id));
    const score = problem ? Number(record.problemScores[problem.id] ?? 0) : 0;
    return { sectionId, sectionLabel: sectionLabels[sectionId], score, maxScore: 20, lostPoints: Math.max(0, 20 - score) };
  });
}

export function buildRecoveryPlan(histories: HistoryEntry[], examSets: ExamSetRecord[]): RecoveryPlan {
  const latestExamSet = [...examSets].sort((a, b) => b.createdAt.localeCompare(a.createdAt))[0];
  const source: RecoveryPlan['source'] = latestExamSet ? '直近90分セット' : '問題別履歴';
  const sectionRows = latestExamSet ? byExamSet(latestExamSet) : byLatestSectionHistories(histories);
  const totalScore = latestExamSet ? latestExamSet.totalScore : sectionRows.reduce((sum, row) => sum + row.score, 0);
  const maxScore = latestExamSet ? 100 : sectionRows.reduce((sum, row) => sum + row.maxScore, 0);
  const biggest = [...sectionRows].sort((a, b) => b.lostPoints - a.lostPoints)[0];
  const missReason = mostFrequentMissReason(histories);
  const cCount = histories.filter((history) => history.rank === 'C').length;
  const sortedLoss = [...sectionRows].filter((row) => row.lostPoints > 0).sort((a, b) => b.lostPoints - a.lostPoints);
  return {
    source,
    totalScore,
    maxScore,
    shortage70: shortage(70, totalScore),
    shortage72: shortage(72, totalScore),
    shortage80: shortage(80, totalScore),
    sectionRows,
    biggestLossSection: biggest ? `${biggest.sectionLabel}（失点${biggest.lostPoints}点）` : '未判定',
    mostFrequentMissReason: missReason,
    nextFocus: focusFor(missReason, cCount),
    recoveryCandidates: sortedLoss.slice(0, 2).map((row) => `${row.sectionLabel}で+${Math.min(row.lostPoints, Math.ceil(shortage(70, totalScore) / Math.max(1, sortedLoss.length)) || row.lostPoints)}点`),
    latestExamSet,
  };
}
