import { useEffect, useMemo, useState } from 'react';
import { AnswerTable, createRows } from './components/AnswerTable';
import { BackupSafetyPanel } from './components/BackupSafetyPanel';
import { DashboardPanel } from './components/DashboardPanel';
import { ExamSetPanel } from './components/ExamSetPanel';
import { ExampleGuide } from './components/ExampleGuide';
import { HistoryPanel } from './components/HistoryPanel';
import { ProblemSelector } from './components/ProblemSelector';
import { ReviewPanel } from './components/ReviewPanel';
import { ScoringPanel } from './components/ScoringPanel';
import { StudyQueuePanel } from './components/StudyQueuePanel';
import { Timer } from './components/Timer';
import { getProblemById, problemCatalog } from './data/problemCatalog';
import { getTemplateById } from './data/templateCatalog';
import { addHistory, clearHistories, deleteHistory, exportAllData, getAllAnswers, getAnswer, getBackupMetadata, getExamSets, getHistories, importAllData, recordBackupMade, saveAnswer, saveExamSet } from './storage/indexedDb';
import type { AnswerState, BackupMetadata, BackupPayload, ExamSetRecord, HistoryEntry, StorageSafetyInfo, TemplateId } from './types';
import { currentAnswerCsvRows, dashboardCsvRows, downloadCsv, downloadText, examSetCsvRows, historyCsvRows, missReasonCsvRows, problemStatsCsvRows, reviewTargetCsvRows, studyQueueCsvRows } from './utils/csv';
import { isDueTodayOrEarlier, nowIso, reviewDateForRank } from './utils/dates';
import { draftSummary, hasMeaningfulAnswer, sumRowPoints } from './utils/scoring';
import { getStorageSafetyInfo } from './utils/storageSafety';
import { buildStudyQueue } from './utils/studyQueue';
import './styles.css';

type HistoryFilter = 'all' | 'today' | 'overdue' | 'c' | 'b';

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

  return {
    ...answer,
    columns,
    rows,
    templateName: template.name,
    updatedAt: nowIso(),
  };
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

export default function App() {
  const [problemId, setProblemId] = useState(problemCatalog[0].id);
  const [answer, setAnswer] = useState<AnswerState>(() => createAnswer(problemCatalog[0].id, problemCatalog[0].defaultTemplateId));
  const [answers, setAnswers] = useState<AnswerState[]>([]);
  const [histories, setHistories] = useState<HistoryEntry[]>([]);
  const [examSets, setExamSets] = useState<ExamSetRecord[]>([]);
  const [, setBackupMeta] = useState<BackupMetadata | null>(null);
  const [safetyInfo, setSafetyInfo] = useState<StorageSafetyInfo | null>(null);
  const [message, setMessage] = useState('待機中');
  const [historyFilter, setHistoryFilter] = useState<HistoryFilter>('all');

  const problem = useMemo(() => getProblemById(problemId), [problemId]);
  const template = useMemo(() => getTemplateById(answer.templateId), [answer.templateId]);
  const studyQueue = useMemo(() => buildStudyQueue(histories), [histories]);

  async function refreshAnswers() {
    setAnswers((await getAllAnswers()).map(normalizeAnswer));
  }

  async function loadHistories() {
    setHistories(await getHistories());
  }

  async function loadExamSets() {
    setExamSets(await getExamSets());
  }

  async function refreshSafety() {
    const [historyRows, answerRows, meta] = await Promise.all([getHistories(), getAllAnswers(), getBackupMetadata()]);
    setBackupMeta(meta);
    setSafetyInfo(await getStorageSafetyInfo(historyRows.length, answerRows.length, meta));
  }

  async function refreshAll() {
    await Promise.all([refreshAnswers(), loadHistories(), loadExamSets(), refreshSafety()]);
  }

  async function loadProblem(nextProblemId: string) {
    const problemDefinition = getProblemById(nextProblemId);
    const stored = await getAnswer(nextProblemId);
    const nextAnswer = normalizeAnswer(stored ?? createAnswer(nextProblemId, problemDefinition.defaultTemplateId));
    setProblemId(nextProblemId);
    setAnswer(nextAnswer);
  }

  useEffect(() => {
    loadProblem(problemCatalog[0].id);
    refreshAll();
  }, []);

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

  const selectTemplate = (templateId: TemplateId) => {
    const nextTemplate = getTemplateById(templateId);
    const keep = window.confirm('テンプレートを変更すると現在の答案テーブルを新しい初期行に置き換えます。変更しますか？');
    if (!keep) return;
    updateAnswer({
      ...answer,
      templateId,
      templateName: nextTemplate.name,
      columns: nextTemplate.columns,
      rows: createRows(nextTemplate.columns.length, nextTemplate.initialRows),
    });
  };

  const applyRowPoints = () => {
    const total = sumRowPoints(answer.rows);
    updateAnswer({ ...answer, rowPointsTotal: total, score: total });
  };

  const manualSave = async () => {
    await saveAnswer(normalizeAnswer({ ...answer, updatedAt: nowIso() }));
    await refreshAll();
    setMessage('手動保存しました');
  };

  const clearCurrentAnswer = async () => {
    if (!window.confirm('現在問題IDの答案を全消去しますか？履歴は消えません。')) return;
    const cleared = createAnswer(problemId, answer.templateId);
    await saveAnswer(cleared);
    setAnswer(cleared);
    await refreshAll();
    setMessage('現在問題IDの答案を全消去しました');
  };

  const bulkAddAllAnswers = async () => {
    const rows = (await getAllAnswers()).map(normalizeAnswer).filter(hasMeaningfulAnswer);
    if (rows.length === 0) {
      setMessage('履歴追加できる保存済み答案がありません');
      return;
    }
    for (const item of rows) {
      await addHistory(buildHistory(item));
    }
    await refreshAll();
    setMessage(`保存済み答案 ${rows.length}件を一括で履歴追加しました`);
  };

  const confirmScoring = async () => {
    const rowTotal = sumRowPoints(answer.rows);
    const scored = normalizeAnswer({ ...answer, rowPointsTotal: rowTotal, score: rowTotal || answer.score, scored: true, scoredAt: nowIso(), updatedAt: nowIso() });
    if (!hasMeaningfulAnswer(scored)) {
      setMessage('空の答案は履歴に追加しませんでした');
      return;
    }
    await saveAnswer(scored);
    await addHistory(buildHistory(scored));
    setAnswer(scored);
    await refreshAll();
    setMessage('採点確定して履歴に追加しました');
  };

  const restoreHistory = async (entry: HistoryEntry) => {
    if (!window.confirm('現在の画面を上書きして、この履歴を復元しますか？')) return;
    const restored = normalizeAnswer(entry.snapshot);
    await saveAnswer(restored);
    setProblemId(entry.problemId);
    setAnswer(restored);
    await refreshAll();
    setMessage('履歴を復元しました');
  };

  const deleteHistoryEntry = async (id: string) => {
    if (!window.confirm('この履歴を削除しますか？')) return;
    await deleteHistory(id);
    await refreshAll();
    setMessage('履歴を削除しました');
  };

  const clearAllHistories = async () => {
    if (!window.confirm('履歴を全削除しますか？答案データは消えません。')) return;
    await clearHistories();
    await refreshAll();
    setMessage('履歴を全削除しました');
  };

  const exportCurrentCsv = () => downloadCsv('current-answer.csv', currentAnswerCsvRows(answer, problem));
  const exportHistoryCsv = () => downloadCsv('history.csv', historyCsvRows(histories));
  const exportReviewCsv = () => downloadCsv('review-targets.csv', reviewTargetCsvRows(histories.filter((history) => isDueTodayOrEarlier(history.nextReviewDate))));
  const exportProblemStatsCsv = () => downloadCsv('problem-stats.csv', problemStatsCsvRows(histories));
  const exportMissReasonCsv = () => downloadCsv('miss-reasons.csv', missReasonCsvRows(histories));
  const exportStudyQueueCsv = () => downloadCsv('study-queue.csv', studyQueueCsvRows(studyQueue));
  const exportDashboardCsv = () => downloadCsv('dashboard.csv', dashboardCsvRows(histories));
  const exportExamSetCsv = () => downloadCsv('exam-sets.csv', examSetCsvRows(examSets));

  const exportJson = async () => {
    const [historyRows, answerRows] = await Promise.all([getHistories(), getAllAnswers()]);
    const meta = await recordBackupMade(historyRows.length, answerRows.length);
    const payload = await exportAllData();
    downloadText('cpa-boki2-backup.json', JSON.stringify({ ...payload, backupMetadata: meta }, null, 2), 'application/json;charset=utf-8');
    setBackupMeta(meta);
    await refreshSafety();
    setMessage('JSONバックアップを出力しました');
  };

  const importJson = async (file: File | undefined) => {
    if (!file) return;
    try {
      const payload = JSON.parse(await file.text()) as BackupPayload;
      await importAllData(payload);
      await loadProblem(problemId);
      await refreshAll();
      setMessage('JSONバックアップを復元しました');
    } catch {
      setMessage('JSONの形式が不正です');
    }
  };

  const saveExamSetRecord = async (record: ExamSetRecord) => {
    const relatedHistoryIds: string[] = [];
    for (const problemState of record.problemStates) {
      if (problemState.rank !== 'B' && problemState.rank !== 'C') continue;
      const problemDefinition = getProblemById(problemState.problemId);
      const base = createAnswer(problemState.problemId, problemDefinition.defaultTemplateId);
      const scored = normalizeAnswer({
        ...base,
        score: Number(problemState.score) || 0,
        rowPointsTotal: Number(problemState.score) || 0,
        rank: problemState.rank,
        nextReviewDate: reviewDateForRank(problemState.rank),
        reviewMemo: `90分セット演習「${record.name}」から登録。${record.memo}`,
        scored: true,
        scoredAt: nowIso(),
        updatedAt: nowIso(),
      });
      const history = buildHistory(scored);
      relatedHistoryIds.push(history.id);
      await addHistory(history);
    }

    await saveExamSet({ ...record, relatedHistoryIds });
    await refreshAll();
    setMessage('90分セット演習履歴を保存しました。B/C判定の問題は復習キューにも反映しました。');
  };

  return (
    <main className="app-shell">
      <header className="app-header">
        <div>
          <h1>CPA日商簿記2級 試験対策編 解答・復習管理アプリ</h1>
          <p>このアプリは自動採点ではありません。問題文・解答・解説はアプリ内に保存しません。紙教材・PDF・CPA画面を見ながら答案を入力し、解答解説を見ながら自己採点してください。</p>
        </div>
        <Timer />
      </header>

      <div className="status-bar">{message}</div>

      <div className="top-actions">
        <button className="danger" onClick={clearCurrentAnswer}>現在問題IDの答案を全消去</button>
        <button onClick={bulkAddAllAnswers}>保存済み答案を一括履歴追加</button>
        <button onClick={exportCurrentCsv}>現在問題IDの答案CSV</button>
        <button onClick={exportHistoryCsv}>履歴一覧CSV</button>
        <button onClick={exportReviewCsv}>復習対象CSV</button>
        <button onClick={exportProblemStatsCsv}>問題ID別成績CSV</button>
        <button onClick={exportMissReasonCsv}>ミス原因別集計CSV</button>
        <button onClick={exportJson}>JSONバックアップ</button>
        <label className="import-label">JSON復元<input type="file" accept="application/json" onChange={(event) => importJson(event.target.files?.[0])} /></label>
        <button onClick={() => window.print()}>印刷</button>
      </div>

      <section className="management-stack">
        <StudyQueuePanel items={studyQueue} onStart={loadProblem} onRestoreLatest={(item) => item.latestHistory && restoreHistory(item.latestHistory)} onExportCsv={exportStudyQueueCsv} />
        <BackupSafetyPanel info={safetyInfo} onBackup={exportJson} onRestoreFile={importJson} onRefresh={refreshSafety} onMessage={setMessage} />
        <DashboardPanel histories={histories} answerCount={answers.length} onExportCsv={exportDashboardCsv} />
        <ExamSetPanel histories={histories} examSets={examSets} onSave={saveExamSetRecord} onOpenProblem={loadProblem} onExportCsv={exportExamSetCsv} />
      </section>

      <div className="main-grid">
        <ProblemSelector selectedProblem={problem} selectedTemplateId={answer.templateId} onSelectProblem={loadProblem} onSelectTemplate={selectTemplate} />

        <section className="center-column">
          <section className="panel problem-info">
            <h2>{problem.sectionLabel} {problem.displayId}</h2>
            <p><strong>{problem.topic}</strong></p>
            <p>使用テンプレート：{answer.templateName}</p>
          </section>
          <ExampleGuide template={template} />
          <AnswerTable answer={answer} template={template} onChange={updateAnswer} onApplyRowPoints={applyRowPoints} />
          <HistoryPanel histories={histories} filter={historyFilter} onFilter={setHistoryFilter} onRestore={restoreHistory} onDelete={deleteHistoryEntry} onClear={clearAllHistories} />
        </section>

        <aside className="right-column">
          <ScoringPanel answer={answer} onChange={updateAnswer} onSave={manualSave} onConfirmScoring={confirmScoring} />
          <ReviewPanel histories={histories} />
        </aside>
      </div>
    </main>
  );
}
