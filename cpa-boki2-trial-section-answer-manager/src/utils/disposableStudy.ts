import { problemCatalog } from '../data/problemCatalog';
import type { HistoryEntry, MasteryMapItem, PracticeSession, StudyQueueItem, TimeMode } from '../types';

export const timeModes: TimeMode[] = ['1分', '3分', '5分', '10分', '30分', '90分'];

export function timeModeDescription(mode: TimeMode): string {
  if (mode === '1分') return '白紙再現カード、ミス原因、同じミスを防ぐ一言を確認';
  if (mode === '3分') return '第1問仕訳1問、短い金額問題、短い復習問題';
  if (mode === '5分') return 'C判定、期限超過、小問を1つ処理';
  if (mode === '10分') return '工業簿記単問、表問題の一部、1論点演習';
  if (mode === '30分') return '大問1つ、重点論点1セット';
  return '第1問〜第5問セット演習';
}

export function latestUnfinishedSession(sessions: PracticeSession[]): PracticeSession | undefined {
  return [...sessions].filter((session) => !session.completed).sort((a, b) => b.startedAt.localeCompare(a.startedAt))[0];
}

export function chooseQuickStartProblem(params: {
  sessions: PracticeSession[];
  studyQueue: StudyQueueItem[];
  masteryMap: MasteryMapItem[];
}): { problemId: string; mode: TimeMode; reason: string } {
  const unfinished = latestUnfinishedSession(params.sessions);
  if (unfinished) return { problemId: unfinished.problemId, mode: unfinished.selectedTimeMode, reason: '未完了の演習セッションを再開' };
  const queueTop = params.studyQueue[0];
  if (queueTop) return { problemId: queueTop.problem.id, mode: '10分', reason: '今日の学習キュー最優先問題' };
  const fallback = params.masteryMap.find((item) => item.status === '危険問題')
    ?? params.masteryMap.find((item) => item.status === '未着手')
    ?? params.masteryMap.find((item) => item.latestRank === 'B')
    ?? params.masteryMap[0];
  return { problemId: fallback?.problem.id ?? problemCatalog[0].id, mode: '10分', reason: '合格到達マップから候補提示' };
}

export function chooseProblemForTimeMode(mode: TimeMode, studyQueue: StudyQueueItem[], masteryMap: MasteryMapItem[], histories: HistoryEntry[]): string {
  const sectionForMode = mode === '3分' ? 'q1' : mode === '10分' ? 'q4' : undefined;
  const queueCandidate = studyQueue.find((item) => !sectionForMode || item.problem.sectionId === sectionForMode) ?? studyQueue[0];
  if (queueCandidate) return queueCandidate.problem.id;

  const notMastered = masteryMap.find((item) => item.status === '期限超過' || item.status === '危険問題' || item.status === '未着手' || item.latestRank === 'B');
  if (notMastered) return notMastered.problem.id;

  const latestByProblem = new Map<string, string>();
  histories.forEach((history) => latestByProblem.set(history.problemId, history.savedAt));
  const oldMastered = [...masteryMap].sort((a, b) => (latestByProblem.get(a.problem.id) ?? '').localeCompare(latestByProblem.get(b.problem.id) ?? ''))[0];
  return oldMastered?.problem.id ?? problemCatalog[0].id;
}
