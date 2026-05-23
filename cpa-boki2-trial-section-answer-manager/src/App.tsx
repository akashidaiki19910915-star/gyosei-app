import { useEffect, useMemo, useState } from 'react';
import { AnswerTable, createRows } from './components/AnswerTable';
import { HistoryPanel } from './components/HistoryPanel';
import { ProblemSelector } from './components/ProblemSelector';
import { ReviewPanel } from './components/ReviewPanel';
import { ScoringPanel } from './components/ScoringPanel';
import { Timer } from './components/Timer';
import { getProblemById, problemCatalog } from './data/problemCatalog';
import { getTemplateById } from './data/templateCatalog';
import { addHistory, clearHistories, deleteHistory, exportAllData, getAllAnswers, getAnswer, getHistories, importAllData, saveAnswer } from './storage/indexedDb';
import type { AnswerState, BackupPayload, HistoryEntry, TemplateId } from './types';
import { currentAnswerCsvRows, downloadCsv, downloadText, historyCsvRows, missReasonCsvRows, problemStatsCsvRows, reviewTargetCsvRows } from './utils/csv';
import { isDueTodayOrEarlier, nowIso } from './utils/dates';
import { draftSummary, hasMeaningfulAnswer, sumRowPoints } from './utils/scoring';
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

function buildHistory(answer: AnswerState): HistoryEntry {
  const problem = getProblemById(answer.problemId);
  const savedAt = nowIso();
  return {
    id: crypto.randomUUID(),
    savedAt,
    subject: problem.subject,
    sectionId: problem.sectionId,
    sectionLabel: problem.sectionLabel,
    problemId: problem.id,
    displayId: problem.displayId,
    topic: problem.topic,
    templateId: answer.templateId,
    templateName: answer.templateName,
    score: answer.score,
    maxScore: answer.maxScore,
    rowPointsTotal: answer.rowPointsTotal,
    scored: answer.scored,
    scoredAt: answer.scoredAt,
    rank: answer.rank,
    nextReviewDate: answer.nextReviewDate,
    missReasons: answer.missReasons,
    reviewMemo: answer.reviewMemo,
    draftMemoSummary: draftSummary(answer.draftMemo),
    snapshot: answer,
  };
}

export default function App() {
  const [problemId, setProblemId] = useState(problemCatalog[0].id);
  const [answer, setAnswer] = useState<AnswerState>(() => createAnswer(problemCatalog[0].id, problemCatalog[0].defaultTemplateId));
  const [histories, setHistories] = useState<HistoryEntry[]>([]);
  const [message, setMessage] = useState('待機中');
  const [historyFilter, setHistoryFilter] = useState<HistoryFilter>('all');

  const problem = useMemo(() => getProblemById(problemId), [problemId]);
  const template = useMemo(() => getTemplateById(answer.templateId), [answer.templateId]);

  async function loadHistories() {
    setHistories(await getHistories());
  }

  async function loadProblem(nextProblemId: string) {
    const problemDefinition = getProblemById(nextProblemId);
    const stored = await getAnswer(nextProblemId);
    setProblemId(nextProblemId);
    setAnswer(stored ?? createAnswer(nextProblemId, problemDefinition.defaultTemplateId));
  }

  useEffect(() => {
    loadProblem(problemCatalog[0].id);
    loadHistories();
  }, []);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      const updated = { ...answer, updatedAt: nowIso() };
      saveAnswer(updated).then(() => setMessage(`自動保存済み ${new Date().toLocaleTimeString('ja-JP')}`));
    }, 800);
    return () => window.clearTimeout(timer);
  }, [answer]);

  const updateAnswer = (next: AnswerState) => setAnswer({ ...next, updatedAt: nowIso() });

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
    await saveAnswer({ ...answer, updatedAt: nowIso() });
    setMessage('手動保存しました');
  };

  const bulkAddAllAnswers = async () => {
    const answers = (await getAllAnswers()).filter(hasMeaningfulAnswer);
    if (answers.length === 0) {
      setMessage('履歴追加できる保存済み答案がありません');
      return;
    }
    for (const item of answers) {
      await addHistory(buildHistory(item));
    }
    await loadHistories();
    setMessage(`保存済み答案 ${answers.length}件を一括で履歴追加しました`);
  };

  const confirmScoring = async () => {
    const rowTotal = sumRowPoints(answer.rows);
    const scored = { ...answer, rowPointsTotal: rowTotal, score: rowTotal || answer.score, scored: true, scoredAt: nowIso(), updatedAt: nowIso() };
    if (!hasMeaningfulAnswer(scored)) {
      setMessage('空の答案は履歴に追加しませんでした');
      return;
    }
    await saveAnswer(scored);
    await addHistory(buildHistory(scored));
    setAnswer(scored);
    await loadHistories();
    setMessage('採点確定して履歴に追加しました');
  };

  const restoreHistory = async (entry: HistoryEntry) => {
    if (!window.confirm('現在の画面を上書きして、この履歴を復元しますか？')) return;
    await saveAnswer(entry.snapshot);
    setProblemId(entry.problemId);
    setAnswer(entry.snapshot);
    setMessage('履歴を復元しました');
  };

  const deleteHistoryEntry = async (id: string) => {
    if (!window.confirm('この履歴を削除しますか？')) return;
    await deleteHistory(id);
    await loadHistories();
    setMessage('履歴を削除しました');
  };

  const clearAllHistories = async () => {
    if (!window.confirm('履歴を全削除しますか？答案データは消えません。')) return;
    await clearHistories();
    await loadHistories();
    setMessage('履歴を全削除しました');
  };

  const exportCurrentCsv = () => downloadCsv('current-answer.csv', currentAnswerCsvRows(answer, problem));
  const exportHistoryCsv = () => downloadCsv('history.csv', historyCsvRows(histories));
  const exportReviewCsv = () => downloadCsv('review-targets.csv', reviewTargetCsvRows(histories.filter((history) => isDueTodayOrEarlier(history.nextReviewDate))));
  const exportProblemStatsCsv = () => downloadCsv('problem-stats.csv', problemStatsCsvRows(histories));
  const exportMissReasonCsv = () => downloadCsv('miss-reasons.csv', missReasonCsvRows(histories));

  const exportJson = async () => {
    const payload = await exportAllData();
    downloadText('cpa-boki2-backup.json', JSON.stringify(payload, null, 2), 'application/json;charset=utf-8');
  };

  const importJson = async (file: File | undefined) => {
    if (!file) return;
    try {
      const payload = JSON.parse(await file.text()) as BackupPayload;
      await importAllData(payload);
      await loadProblem(problemId);
      await loadHistories();
      setMessage('JSONバックアップを復元しました');
    } catch {
      setMessage('JSONの形式が不正です');
    }
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

      <div className="main-grid">
        <ProblemSelector selectedProblem={problem} selectedTemplateId={answer.templateId} onSelectProblem={loadProblem} onSelectTemplate={selectTemplate} />

        <section className="center-column">
          <section className="panel problem-info">
            <h2>{problem.sectionLabel} {problem.displayId}</h2>
            <p><strong>{problem.topic}</strong></p>
            <p>使用テンプレート：{answer.templateName}</p>
          </section>
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
