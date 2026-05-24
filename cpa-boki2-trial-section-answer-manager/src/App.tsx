import { useEffect, useMemo, useState } from 'react';
import { createRows } from './components/AnswerTable';
import { BackupSafetyPanel } from './components/BackupSafetyPanel';
import { DashboardPanel } from './components/DashboardPanel';
import { ExamSetPanel } from './components/ExamSetPanel';
import { FocusedGradingPanel } from './components/FocusedGradingPanel';
import { FocusedPracticePanel } from './components/FocusedPracticePanel';
import { HistoryPanel } from './components/HistoryPanel';
import { MasteryMapPanel } from './components/MasteryMapPanel';
import { MaterialSetupWorkspace } from './components/MaterialSetupWorkspace';
import { MistakeCardPanel } from './components/MistakeCardPanel';
import { RecoveryPlanPanel } from './components/RecoveryPlanPanel';
import { ReviewPanel } from './components/ReviewPanel';
import { StudyQueuePanel } from './components/StudyQueuePanel';
import { Timer } from './components/Timer';
import { getProblemById, problemCatalog } from './data/problemCatalog';
import { getTemplateById } from './data/templateCatalog';
import {
  addHistory,
  clearHistories,
  deleteHistory,
  deleteMaterialPdf,
  deleteMaterialPdfMapping,
  deleteMistakeCard,
  exportAllData,
  getAllAnswers,
  getAnswer,
  getBackupMetadata,
  getExamSets,
  getHistories,
  getMaterialPdfMappings,
  getMaterialPdfs,
  getMistakeCards,
  getPracticeSessions,
  getQuestionCards,
  getReviewStates,
  importAllData,
  recordBackupMade,
  saveAnswer,
  saveExamSet,
  saveMaterialPdf,
  saveMaterialPdfMapping,
  saveMistakeCard,
  savePracticeSession,
  saveQuestionCard,
  saveReviewState,
} from './storage/indexedDb';
import type { AnswerState, BackupMetadata, BackupPayload, ExamSetRecord, GradingStatus, HistoryEntry, MaterialPdf, MaterialPdfMapping, MistakeCard, PracticeSession, QuestionCard, ReviewState, StorageSafetyInfo, TemplateId, TimeMode } from './types';
import { currentAnswerCsvRows, dashboardCsvRows, downloadCsv, downloadText, examSetCsvRows, historyCsvRows, masteryMapCsvRows, missReasonCsvRows, mistakeCardCsvRows, problemStatsCsvRows, recoveryPlanCsvRows, reviewTargetCsvRows, studyQueueCsvRows } from './utils/csv';
import { chooseQuickStartProblem } from './utils/disposableStudy';
import { isDueTodayOrEarlier, nowIso, reviewDateForRank } from './utils/dates';
import { buildMasteryMap } from './utils/mastery';
import { fallbackQuestionCard, questionCardFromMapping } from './utils/questionCards';
import { buildRecoveryPlan } from './utils/recovery';
import { rankFromSelfGrading, updateReviewStateFromGrading } from './utils/reviewScheduler';
import { draftSummary, hasMeaningfulAnswer, sumRowPoints } from './utils/scoring';
import { getStorageSafetyInfo } from './utils/storageSafety';
import { buildStudyQueue } from './utils/studyQueue';
import './styles.css';

type HistoryFilter = 'all' | 'today' | 'overdue' | 'c' | 'b';
type AppMode = 'study' | 'grading' | 'materials' | 'progress';

function createAnswer(problemId: string, templateId: TemplateId): AnswerState {
  const template = getTemplateById(templateId);
  const now = nowIso();
  return {
    id: problemId,
    problemId,
    templateId,
    templateName: template.name,
    columns: template.columns,
    rows: createRows(template.columns.length, template.initialRows),
    problemBlocks: [],
    draftMemo: '',
    reviewMemo: '',
    score: 0,
    maxScore: 20,
    rowPointsTotal: 0,
    rank: '',
    nextReviewDate: '',
    missReasons: [],
    scoringStarted: false,
    scored: false,
    scoredAt: '',
    updatedAt: now,
    createdAt: now,
  };
}

function normalizeAnswer(answer: AnswerState): AnswerState {
  const template = getTemplateById(answer.templateId);
  if (answer.templateId !== 'journal') return answer;
  const legacyIndex = answer.columns.indexOf('会社名・立場');
  const columns = template.columns;
  if (legacyIndex === -1 && answer.columns.join('|') === columns.join('|')) return answer;
  const rows = answer.rows.map((row) => {
    const cells = [...row.cells];
    if (legacyIndex >= 0) cells.splice(legacyIndex, 1);
    return { ...row, cells: columns.map((_, index) => cells[index] ?? '') };
  });
  return { ...answer, columns, rows, templateName: template.name, updatedAt: nowIso() };
}

function buildHistory(answer: AnswerState): HistoryEntry {
  const problem = getProblemById(answer.problemId);
  const savedAt = nowIso();
  const snapshot = normalizeAnswer(answer);
  return {
    id: crypto.randomUUID(),
    savedAt,
    subject: problem.subject,
    sectionId: problem.sectionId,
    sectionLabel: problem.sectionLabel,
    problemId: problem.id,
    displayId: problem.displayId,
    topic: problem.topic,
    templateId: snapshot.templateId,
    templateName: snapshot.templateName,
    score: snapshot.score,
    maxScore: snapshot.maxScore,
    rowPointsTotal: snapshot.rowPointsTotal,
    scored: snapshot.scored,
    scoredAt: snapshot.scoredAt,
    rank: snapshot.rank,
    nextReviewDate: snapshot.nextReviewDate,
    missReasons: snapshot.missReasons,
    reviewMemo: snapshot.reviewMemo,
    draftMemoSummary: draftSummary(snapshot.draftMemo),
    snapshot,
  };
}

function secondsBetween(start: string, end: string): number {
  const startMs = new Date(start).getTime();
  const endMs = new Date(end).getTime();
  if (!Number.isFinite(startMs) || !Number.isFinite(endMs)) return 0;
  return Math.max(0, Math.round((endMs - startMs) / 1000));
}

export default function App() {
  const [problemId, setProblemId] = useState(problemCatalog[0].id);
  const [answer, setAnswer] = useState<AnswerState>(() => createAnswer(problemCatalog[0].id, problemCatalog[0].defaultTemplateId));
  const [answers, setAnswers] = useState<AnswerState[]>([]);
  const [histories, setHistories] = useState<HistoryEntry[]>([]);
  const [examSets, setExamSets] = useState<ExamSetRecord[]>([]);
  const [mistakeCards, setMistakeCards] = useState<MistakeCard[]>([]);
  const [materialPdfs, setMaterialPdfs] = useState<MaterialPdf[]>([]);
  const [pdfMappings, setPdfMappings] = useState<MaterialPdfMapping[]>([]);
  const [questionCards, setQuestionCards] = useState<QuestionCard[]>([]);
  const [practiceSessions, setPracticeSessions] = useState<PracticeSession[]>([]);
  const [reviewStates, setReviewStates] = useState<ReviewState[]>([]);
  const [activeSessionId, setActiveSessionId] = useState('');
  const [, setBackupMeta] = useState<BackupMetadata | null>(null);
  const [safetyInfo, setSafetyInfo] = useState<StorageSafetyInfo | null>(null);
  const [message, setMessage] = useState('待機中');
  const [historyFilter, setHistoryFilter] = useState<HistoryFilter>('all');
  const [mode, setMode] = useState<AppMode>('study');

  const problem = useMemo(() => getProblemById(problemId), [problemId]);
  const template = useMemo(() => getTemplateById(answer.templateId), [answer.templateId]);
  const currentMapping = useMemo(() => pdfMappings.find((item) => item.problemId === problemId), [pdfMappings, problemId]);
  const currentQuestionCard = useMemo(() => questionCards.find((card) => card.sourceProblemId === problemId || card.id === problemId) ?? (currentMapping ? questionCardFromMapping(currentMapping) : fallbackQuestionCard(problem)), [questionCards, currentMapping, problem, problemId]);
  const currentPdf = useMemo(() => materialPdfs.find((item) => item.id === currentQuestionCard.sourceMaterialId || item.id === currentMapping?.materialPdfId), [materialPdfs, currentQuestionCard.sourceMaterialId, currentMapping?.materialPdfId]);
  const studyQueue = useMemo(() => buildStudyQueue(histories), [histories]);
  const masteryMap = useMemo(() => buildMasteryMap(histories), [histories]);
  const recoveryPlan = useMemo(() => buildRecoveryPlan(histories, examSets), [histories, examSets]);
  const activeSession = useMemo(() => practiceSessions.find((session) => session.id === activeSessionId) ?? null, [practiceSessions, activeSessionId]);
  const currentReviewState = useMemo(() => reviewStates.find((state) => state.problemId === problemId), [reviewStates, problemId]);
  const quickChoice = useMemo(() => chooseQuickStartProblem({ sessions: practiceSessions, studyQueue, masteryMap }), [practiceSessions, studyQueue, masteryMap]);
  const progressIndex = useMemo(() => Math.max(1, studyQueue.findIndex((item) => item.problem.id === problemId) + 1 || 1), [studyQueue, problemId]);
  const progressSummary = useMemo(() => ({
    dueToday: studyQueue.filter((item) => item.statusLabels.includes('今日復習')).length,
    overdue: studyQueue.filter((item) => item.statusLabels.includes('期限超過')).length,
    cRank: masteryMap.filter((item) => item.latestRank === 'C' || item.status === '危険問題').length,
    untouched: masteryMap.filter((item) => item.status === '未着手').length,
    last7Days: histories.filter((history) => Date.now() - new Date(history.savedAt).getTime() <= 7 * 24 * 60 * 60 * 1000).length,
    last30Days: histories.filter((history) => Date.now() - new Date(history.savedAt).getTime() <= 30 * 24 * 60 * 60 * 1000).length,
  }), [studyQueue, masteryMap, histories]);

  async function refreshAnswers() { setAnswers((await getAllAnswers()).map(normalizeAnswer)); }
  async function loadHistories() { setHistories(await getHistories()); }
  async function loadExamSets() { setExamSets(await getExamSets()); }
  async function loadMistakeCards() { setMistakeCards(await getMistakeCards()); }
  async function loadMaterialPdfs() { setMaterialPdfs(await getMaterialPdfs()); }
  async function loadPdfMappings() { setPdfMappings(await getMaterialPdfMappings()); }
  async function loadQuestionCards() { setQuestionCards(await getQuestionCards()); }
  async function loadPracticeSessions() { setPracticeSessions(await getPracticeSessions()); }
  async function loadReviewStates() { setReviewStates(await getReviewStates()); }

  async function refreshSafety() {
    const [historyRows, answerRows, meta] = await Promise.all([getHistories(), getAllAnswers(), getBackupMetadata()]);
    setBackupMeta(meta);
    setSafetyInfo(await getStorageSafetyInfo(historyRows.length, answerRows.length, meta));
  }

  async function refreshAll() {
    await Promise.all([refreshAnswers(), loadHistories(), loadExamSets(), loadMistakeCards(), loadMaterialPdfs(), loadPdfMappings(), loadQuestionCards(), loadPracticeSessions(), loadReviewStates(), refreshSafety()]);
  }

  async function loadProblem(nextProblemId: string) {
    const problemDefinition = getProblemById(nextProblemId);
    const stored = await getAnswer(nextProblemId);
    const nextAnswer = normalizeAnswer(stored ?? createAnswer(nextProblemId, problemDefinition.defaultTemplateId));
    setProblemId(nextProblemId);
    setAnswer(nextAnswer);
  }

  useEffect(() => { loadProblem(problemCatalog[0].id); refreshAll(); }, []);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      const updated = normalizeAnswer({ ...answer, updatedAt: nowIso() });
      saveAnswer(updated).then(() => {
        setMessage(`自動保存済み ${new Date().toLocaleTimeString('ja-JP')}`);
        refreshAnswers();
      });
    }, 800);
    return () => window.clearTimeout(timer);
  }, [answer]);

  const updateAnswer = (next: AnswerState) => setAnswer(normalizeAnswer({ ...next, updatedAt: nowIso() }));

  async function createPracticeSession(problemIdForSession: string, selectedTimeMode: TimeMode): Promise<PracticeSession> {
    const problemDefinition = getProblemById(problemIdForSession);
    const stored = await getAnswer(problemIdForSession);
    const baseAnswer = normalizeAnswer(stored ?? createAnswer(problemIdForSession, problemDefinition.defaultTemplateId));
    const now = nowIso();
    return {
      id: crypto.randomUUID(),
      examType: '日商簿記2級',
      problemId: problemIdForSession,
      startedAt: now,
      submittedAt: '',
      durationSeconds: 0,
      selectedTimeMode,
      answerSnapshot: baseAnswer,
      gradingMode: 'self',
      gradingResult: { status: '', score: 0, maxScore: baseAnswer.maxScore, scoreRate: 0, autoGraded: false, detailRows: [] },
      openedAnswer: false,
      openedExplanation: false,
      memo: '',
      completed: false,
    };
  }

  async function startCurrentSession(): Promise<PracticeSession> {
    if (activeSession && activeSession.problemId === problem.id && !activeSession.completed) return activeSession;
    const session = await createPracticeSession(problem.id, currentMapping?.estimatedMinutes ?? '10分');
    await savePracticeSession(session);
    setActiveSessionId(session.id);
    await loadPracticeSessions();
    return session;
  }

  const applyRowPoints = () => {
    const total = sumRowPoints(answer.rows);
    updateAnswer({ ...answer, rowPointsTotal: total, score: total });
  };

  const manualSave = async () => { await saveAnswer(normalizeAnswer({ ...answer, updatedAt: nowIso() })); await refreshAll(); setMessage('手動保存しました'); };
  const clearCurrentAnswer = async () => { if (!window.confirm('現在問題IDの答案を全消去しますか？履歴は消えません。')) return; const cleared = createAnswer(problemId, answer.templateId); await saveAnswer(cleared); setAnswer(cleared); await refreshAll(); setMessage('現在問題IDの答案を全消去しました'); };
  const bulkAddAllAnswers = async () => { const rows = (await getAllAnswers()).map(normalizeAnswer).filter(hasMeaningfulAnswer); if (rows.length === 0) { setMessage('履歴追加できる保存済み答案がありません'); return; } for (const item of rows) await addHistory(buildHistory(item)); await refreshAll(); setMessage(`保存済み答案 ${rows.length}件を一括で履歴追加しました`); };
  const confirmScoring = async () => { const rowTotal = sumRowPoints(answer.rows); const scored = normalizeAnswer({ ...answer, rowPointsTotal: rowTotal, score: rowTotal || answer.score, scored: true, scoredAt: nowIso(), updatedAt: nowIso() }); if (!hasMeaningfulAnswer(scored)) { setMessage('空の答案は履歴に追加しませんでした'); return; } await saveAnswer(scored); await addHistory(buildHistory(scored)); setAnswer(scored); await refreshAll(); setMessage('採点確定して履歴に追加しました'); };

  async function submitSelfGrading(input: { status: GradingStatus; score: number; maxScore: number; memo: string }) {
    const submittedAt = nowIso();
    const rank = rankFromSelfGrading(input.status, input.score, input.maxScore);
    const reviewState = updateReviewStateFromGrading({ problemId: problem.id, previous: currentReviewState, status: input.status, score: input.score, maxScore: input.maxScore });
    const scored = normalizeAnswer({ ...answer, score: input.score, maxScore: input.maxScore, scored: true, scoringStarted: true, scoredAt: submittedAt, rank, nextReviewDate: reviewState.nextReviewDate, reviewMemo: input.memo || answer.reviewMemo, updatedAt: submittedAt });
    const session = activeSession && activeSession.problemId === problem.id && !activeSession.completed ? activeSession : await startCurrentSession();
    const completedSession: PracticeSession = { ...session, submittedAt, durationSeconds: secondsBetween(session.startedAt, submittedAt), answerSnapshot: scored, gradingResult: { status: input.status, score: input.score, maxScore: input.maxScore, scoreRate: input.maxScore > 0 ? Math.round((input.score / input.maxScore) * 1000) / 10 : 0, autoGraded: false, detailRows: [] }, reviewState, openedAnswer: true, openedExplanation: Boolean(currentMapping?.answerText || currentMapping?.explanationText), memo: input.memo, completed: true };
    await saveAnswer(scored);
    await addHistory(buildHistory(scored));
    await saveReviewState(reviewState);
    await savePracticeSession(completedSession);
    setAnswer(scored);
    setActiveSessionId('');
    await refreshAll();
    setMessage(`自己採点を保存しました。判定 ${rank}、次回復習日 ${reviewState.nextReviewDate}`);
    await openNextProblem();
  }

  const restoreHistory = async (entry: HistoryEntry) => { if (!window.confirm('現在の画面を上書きして、この履歴を復元しますか？')) return; const restored = normalizeAnswer(entry.snapshot); await saveAnswer(restored); setProblemId(entry.problemId); setAnswer(restored); setMode('study'); await refreshAll(); setMessage('履歴を復元しました'); };
  const deleteHistoryEntry = async (id: string) => { if (!window.confirm('この履歴を削除しますか？')) return; await deleteHistory(id); await refreshAll(); setMessage('履歴を削除しました'); };
  const clearAllHistories = async () => { if (!window.confirm('履歴を全削除しますか？答案データは消えません。')) return; await clearHistories(); await refreshAll(); setMessage('履歴を全削除しました'); };
  const saveCard = async (card: MistakeCard) => { await saveMistakeCard(card); await loadMistakeCards(); setMessage('白紙再現・解き直しカードを保存しました'); };
  const removeCard = async (id: string) => { if (!window.confirm('この解き直しカードを削除しますか？')) return; await deleteMistakeCard(id); await loadMistakeCards(); setMessage('白紙再現・解き直しカードを削除しました'); };
  const savePdf = async (pdf: MaterialPdf) => { await saveMaterialPdf(pdf); await loadMaterialPdfs(); };
  const removePdf = async (id: string) => { await deleteMaterialPdf(id); await Promise.all([loadMaterialPdfs(), loadPdfMappings(), loadQuestionCards()]); setMessage('PDF教材データと紐付けを削除しました。答案履歴・白紙再現カードは保持されています。'); };
  const savePdfMapping = async (mapping: MaterialPdfMapping) => { await saveMaterialPdfMapping(mapping); await saveQuestionCard(questionCardFromMapping(mapping)); await Promise.all([loadPdfMappings(), loadQuestionCards()]); setMessage('問題カードを保存しました'); };
  const removePdfMapping = async (id: string) => { if (!window.confirm('このPDF抽出テキスト紐づけを削除しますか？PDF教材データと答案履歴は削除されません。')) return; await deleteMaterialPdfMapping(id); await Promise.all([loadPdfMappings(), loadQuestionCards()]); setMessage('PDF抽出テキスト紐づけを削除しました'); };

  const exportCurrentCsv = () => downloadCsv('current-answer.csv', currentAnswerCsvRows(answer, problem));
  const exportHistoryCsv = () => downloadCsv('history.csv', historyCsvRows(histories));
  const exportReviewCsv = () => downloadCsv('review-targets.csv', reviewTargetCsvRows(histories.filter((history) => isDueTodayOrEarlier(history.nextReviewDate))));
  const exportProblemStatsCsv = () => downloadCsv('problem-stats.csv', problemStatsCsvRows(histories));
  const exportMissReasonCsv = () => downloadCsv('miss-reasons.csv', missReasonCsvRows(histories));
  const exportStudyQueueCsv = () => downloadCsv('study-queue.csv', studyQueueCsvRows(studyQueue));
  const exportDashboardCsv = () => downloadCsv('dashboard.csv', dashboardCsvRows(histories));
  const exportExamSetCsv = () => downloadCsv('exam-sets.csv', examSetCsvRows(examSets));
  const exportMasteryCsv = () => downloadCsv('mastery-map.csv', masteryMapCsvRows(masteryMap));
  const exportMistakeCardsCsv = () => downloadCsv('mistake-cards.csv', mistakeCardCsvRows(mistakeCards, getProblemById));
  const exportRecoveryCsv = () => downloadCsv('recovery-plan.csv', recoveryPlanCsvRows(recoveryPlan));
  const exportJson = async () => { const [historyRows, answerRows] = await Promise.all([getHistories(), getAllAnswers()]); const meta = await recordBackupMade(historyRows.length, answerRows.length); const payload = await exportAllData(); downloadText('cpa-boki2-backup.json', JSON.stringify({ ...payload, backupMetadata: meta }, null, 2), 'application/json;charset=utf-8'); setBackupMeta(meta); await refreshSafety(); setMessage('JSONバックアップを出力しました。PDF本体と抽出教材本文は含めていません。'); };
  const importJson = async (file: File | undefined) => { if (!file) return; try { const payload = JSON.parse(await file.text()) as BackupPayload; await importAllData(payload); await loadProblem(problemId); await refreshAll(); setMessage('JSONバックアップを復元しました。PDF本体・抽出テキストは端末内データを維持する仕様です。'); } catch { setMessage('JSONの形式が不正です'); } };

  const saveExamSetRecord = async (record: ExamSetRecord) => { const relatedHistoryIds: string[] = []; for (const problemState of record.problemStates) { if (problemState.rank !== 'B' && problemState.rank !== 'C') continue; const problemDefinition = getProblemById(problemState.problemId); const base = createAnswer(problemState.problemId, problemDefinition.defaultTemplateId); const scored = normalizeAnswer({ ...base, score: Number(problemState.score) || 0, rowPointsTotal: Number(problemState.score) || 0, rank: problemState.rank, nextReviewDate: reviewDateForRank(problemState.rank), reviewMemo: `90分セット演習「${record.name}」から登録。${record.memo}`, scored: true, scoredAt: nowIso(), updatedAt: nowIso() }); const history = buildHistory(scored); relatedHistoryIds.push(history.id); await addHistory(history); } await saveExamSet({ ...record, relatedHistoryIds }); await refreshAll(); setMessage('90分セット演習履歴を保存しました。B/C判定の問題は復習キューにも反映しました。'); };
  const openNextProblem = async () => { const next = studyQueue[0]?.problem.id ?? quickChoice.problemId; await loadProblem(next); setMode('study'); };

  return (
    <main className="app-shell redesigned-shell">
      <header className="app-header compact-header"><div><h1>CPA日商簿記2級 試験対策編 解答・復習管理アプリ</h1><p>問題文を読む → 答案入力 → 採点 → 次回復習日へ回す、を最短で回すための学習画面です。</p></div><Timer /></header>
      <nav className="mode-nav" aria-label="学習モード切替"><button className={mode === 'study' ? 'active' : ''} onClick={() => setMode('study')}>今すぐ解く</button><button className={mode === 'grading' ? 'active' : ''} onClick={() => setMode('grading')}>採点する</button><button className={mode === 'materials' ? 'active' : ''} onClick={() => setMode('materials')}>教材を設定する</button><button className={mode === 'progress' ? 'active' : ''} onClick={() => setMode('progress')}>復習を見る</button></nav>
      <div className="status-bar">{message}</div>

      {mode === 'study' && <section className="mode-section study-mode-section">
        {currentQuestionCard.needsReview && <section className="panel first-use-steps"><h2>問題カードの確認が必要です</h2><div><span>1. 教材を設定する</span><span>2. 抽出結果を整形する</span><span>3. 今すぐ解く</span></div><button className="accent" onClick={() => setMode('materials')}>教材を設定する</button></section>}
        <FocusedPracticePanel problem={problem} questionCard={currentQuestionCard} answer={answer} template={template} mapping={currentMapping} pdf={currentPdf} totalToday={Math.max(1, studyQueue.length || 1)} progressIndex={progressIndex} onChange={updateAnswer} onSave={manualSave} onGoGrading={() => setMode('grading')} onNextProblem={openNextProblem} onApplyRowPoints={applyRowPoints} />
      </section>}

      {mode === 'grading' && <FocusedGradingPanel problem={problem} answer={answer} template={template} mapping={currentMapping} pdf={currentPdf} onChange={updateAnswer} onSave={manualSave} onApplyRowPoints={applyRowPoints} onSubmitSelfGrading={submitSelfGrading} onBackPractice={() => setMode('study')} />}

      {mode === 'materials' && <MaterialSetupWorkspace pdfs={materialPdfs} mappings={pdfMappings} selectedProblemId={problemId} onSavePdf={savePdf} onDeletePdf={removePdf} onSaveMapping={savePdfMapping} onDeleteMapping={removePdfMapping} onOpenProblem={loadProblem} onMessage={setMessage} />}

      {mode === 'progress' && <section className="mode-section progress-mode-section"><section className="panel mode-card progress-overview-panel"><div><p className="eyebrow">復習を見る</p><h2>今日やる問題だけ先に見る</h2><p>詳細な管理機能は下の折りたたみの中に退避しています。</p></div><div className="progress-summary-grid"><div><strong>{progressSummary.dueToday}</strong><span>今日やる問題</span></div><div><strong>{progressSummary.overdue}</strong><span>期限超過</span></div><div><strong>{progressSummary.cRank}</strong><span>C判定</span></div><div><strong>{progressSummary.untouched}</strong><span>未着手</span></div><div><strong>{progressSummary.last7Days}</strong><span>直近7日演習</span></div><div><strong>{progressSummary.last30Days}</strong><span>直近30日演習</span></div></div><button className="resume-button" onClick={() => { void openNextProblem(); }}>次に解く</button></section><details className="panel management-panel" open><summary>今日の復習・期限超過・C/B判定</summary><StudyQueuePanel items={studyQueue} onStart={loadProblem} onRestoreLatest={(item) => item.latestHistory && restoreHistory(item.latestHistory)} onExportCsv={exportStudyQueueCsv} /></details><details className="panel management-panel"><summary>合格到達マップ</summary><MasteryMapPanel histories={histories} onOpenProblem={loadProblem} onExportCsv={exportMasteryCsv} /></details><details className="panel management-panel"><summary>白紙再現カード</summary><MistakeCardPanel problem={problem} answer={answer} cards={mistakeCards} onSave={saveCard} onDelete={removeCard} onExportCsv={exportMistakeCardsCsv} /></details><details className="panel management-panel"><summary>70点リカバリー表・90分セット</summary><RecoveryPlanPanel histories={histories} examSets={examSets} onExportCsv={exportRecoveryCsv} /><ExamSetPanel histories={histories} examSets={examSets} onSave={saveExamSetRecord} onOpenProblem={loadProblem} onExportCsv={exportExamSetCsv} /></details><details className="panel management-panel"><summary>履歴一覧・復習管理</summary><ReviewPanel histories={histories} /><HistoryPanel histories={histories} filter={historyFilter} onFilter={setHistoryFilter} onRestore={restoreHistory} onDelete={deleteHistoryEntry} onClear={clearAllHistories} /></details><details className="panel management-panel"><summary>CSV出力・JSONバックアップ・保存状態</summary><div className="top-actions management-actions"><button className="danger" onClick={clearCurrentAnswer}>現在問題IDの答案を全消去</button><button onClick={bulkAddAllAnswers}>保存済み答案を一括履歴追加</button><button onClick={exportCurrentCsv}>現在問題IDの答案CSV</button><button onClick={exportHistoryCsv}>履歴一覧CSV</button><button onClick={exportReviewCsv}>復習対象CSV</button><button onClick={exportProblemStatsCsv}>問題ID別成績CSV</button><button onClick={exportMissReasonCsv}>ミス原因別集計CSV</button><button onClick={exportDashboardCsv}>ダッシュボードCSV</button><button onClick={exportExamSetCsv}>90分セットCSV</button><button onClick={exportMasteryCsv}>合格到達マップCSV</button><button onClick={exportRecoveryCsv}>70点リカバリーCSV</button><button onClick={exportJson}>JSONバックアップ</button><label className="import-label">JSON復元<input type="file" accept="application/json" onChange={(event) => importJson(event.target.files?.[0])} /></label><button onClick={() => window.print()}>印刷</button></div><BackupSafetyPanel info={safetyInfo} onBackup={exportJson} onRestoreFile={importJson} onRefresh={refreshSafety} onMessage={setMessage} /><DashboardPanel histories={histories} answerCount={answers.length} onExportCsv={exportDashboardCsv} /></details></section>}
    </main>
  );
}
