import { useEffect, useMemo, useRef, useState } from 'react';
import { problemCatalog } from '../data/problemCatalog';
import type { ExamSetProblemState, ExamSetRecord, HistoryEntry, ProblemDefinition, ReviewRank, SectionId } from '../types';
import { nowIso } from '../utils/dates';

const sectionOrder: SectionId[] = ['q1', 'q2', 'q3', 'q4', 'q5'];
const EXAM_SECONDS = 90 * 60;

interface Props {
  histories: HistoryEntry[];
  examSets: ExamSetRecord[];
  onSave: (record: ExamSetRecord) => void;
  onOpenProblem: (problemId: string) => void;
  onExportCsv: () => void;
}

function firstBySection(sectionId: SectionId): ProblemDefinition {
  return problemCatalog.find((problem) => problem.sectionId === sectionId)!;
}

function defaultSelected(): string[] {
  return sectionOrder.map((sectionId) => firstBySection(sectionId).id);
}

function latestHistories(histories: HistoryEntry[]) {
  const map = new Map<string, HistoryEntry>();
  histories.forEach((history) => {
    const current = map.get(history.problemId);
    if (!current || history.savedAt > current.savedAt) map.set(history.problemId, history);
  });
  return map;
}

export function ExamSetPanel({ histories, examSets, onSave, onOpenProblem, onExportCsv }: Props) {
  const [name, setName] = useState('90分セット演習');
  const [selectedIds, setSelectedIds] = useState<string[]>(defaultSelected);
  const [problemStates, setProblemStates] = useState<ExamSetProblemState[]>(() => defaultSelected().map((problemId) => ({ problemId, status: '未開始', score: 0, rank: '' })));
  const [memo, setMemo] = useState('');
  const [startedAt, setStartedAt] = useState('');
  const [finishedAt, setFinishedAt] = useState('');
  const [seconds, setSeconds] = useState(EXAM_SECONDS);
  const [running, setRunning] = useState(false);
  const timerRef = useRef<number | null>(null);
  const latest = useMemo(() => latestHistories(histories), [histories]);

  useEffect(() => {
    if (!running) return;
    timerRef.current = window.setInterval(() => setSeconds((value) => Math.max(value - 1, 0)), 1000);
    return () => {
      if (timerRef.current !== null) window.clearInterval(timerRef.current);
      timerRef.current = null;
    };
  }, [running]);

  useEffect(() => {
    if (seconds === 0 && running) setRunning(false);
  }, [seconds, running]);

  const totalScore = problemStates.reduce((sum, state) => sum + (Number(state.score) || 0), 0);
  const minutes = Math.floor(seconds / 60).toString().padStart(2, '0');
  const rest = (seconds % 60).toString().padStart(2, '0');

  const updateProblem = (sectionIndex: number, problemId: string) => {
    const nextIds = [...selectedIds];
    nextIds[sectionIndex] = problemId;
    setSelectedIds(nextIds);
    setProblemStates(nextIds.map((id) => problemStates.find((state) => state.problemId === id) ?? { problemId: id, status: '未開始', score: 0, rank: '' }));
  };

  const updateState = (problemId: string, patch: Partial<ExamSetProblemState>) => {
    setProblemStates((rows) => rows.map((row) => row.problemId === problemId ? { ...row, ...patch } : row));
  };

  const autoUntouched = () => {
    const selected = sectionOrder.map((sectionId) => problemCatalog.find((problem) => problem.sectionId === sectionId && !latest.has(problem.id)) ?? firstBySection(sectionId));
    setSelectedIds(selected.map((problem) => problem.id));
    setProblemStates(selected.map((problem) => ({ problemId: problem.id, status: '未開始', score: 0, rank: '' })));
  };

  const autoC = () => {
    const selected = sectionOrder.map((sectionId) => problemCatalog.find((problem) => latest.get(problem.id)?.rank === 'C' && problem.sectionId === sectionId) ?? firstBySection(sectionId));
    setSelectedIds(selected.map((problem) => problem.id));
    setProblemStates(selected.map((problem) => ({ problemId: problem.id, status: '未開始', score: 0, rank: '' })));
  };

  const autoRandom = () => {
    const selected = sectionOrder.map((sectionId) => {
      const rows = problemCatalog.filter((problem) => problem.sectionId === sectionId);
      return rows[Math.floor(Math.random() * rows.length)] ?? firstBySection(sectionId);
    });
    setSelectedIds(selected.map((problem) => problem.id));
    setProblemStates(selected.map((problem) => ({ problemId: problem.id, status: '未開始', score: 0, rank: '' })));
  };

  const start = () => {
    if (!startedAt) setStartedAt(nowIso());
    setRunning(true);
  };

  const finishAndSave = () => {
    const end = nowIso();
    setFinishedAt(end);
    setRunning(false);
    const record: ExamSetRecord = {
      id: crypto.randomUUID(),
      name,
      createdAt: nowIso(),
      startedAt: startedAt || nowIso(),
      finishedAt: end,
      durationSeconds: EXAM_SECONDS - seconds,
      selectedProblemIds: selectedIds,
      problemStates,
      problemScores: Object.fromEntries(problemStates.map((state) => [state.problemId, Number(state.score) || 0])),
      totalScore,
      passLineReached: totalScore >= 70,
      memo,
      relatedHistoryIds: [],
    };
    onSave(record);
  };

  return (
    <details className="panel management-panel">
      <summary>90分セット演習モード</summary>
      <div className="panel-body">
        <div className="exam-timer"><strong>{minutes}:{rest}</strong><span>{running ? '演習中' : seconds === 0 ? '終了' : '待機中'}</span></div>
        <div className="button-row"><button onClick={start}>90分タイマー開始</button><button className="secondary" onClick={() => setRunning(false)}>一時停止</button><button className="secondary" onClick={() => { setRunning(false); setSeconds(EXAM_SECONDS); setStartedAt(''); setFinishedAt(''); }}>リセット</button><button onClick={autoUntouched}>未着手優先で自動作成</button><button onClick={autoC}>C判定優先で自動作成</button><button onClick={autoRandom}>ランダムセット作成</button><button onClick={onExportCsv}>セット履歴CSV</button></div>
        <label>セット名<input value={name} onChange={(event) => setName(event.target.value)} /></label>
        <div className="table-wrap compact-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>大問</th><th>問題ID</th><th>状態</th><th>得点</th><th>判定</th><th>操作</th></tr></thead>
            <tbody>
              {sectionOrder.map((sectionId, index) => {
                const problems = problemCatalog.filter((problem) => problem.sectionId === sectionId);
                const selected = selectedIds[index];
                const state = problemStates.find((row) => row.problemId === selected) ?? { problemId: selected, status: '未開始', score: 0, rank: '' as ReviewRank };
                return (
                  <tr key={sectionId}>
                    <td>{firstBySection(sectionId).sectionLabel}</td>
                    <td><select value={selected} onChange={(event) => updateProblem(index, event.target.value)}>{problems.map((problem) => <option key={problem.id} value={problem.id}>{problem.displayId} {problem.topic}</option>)}</select></td>
                    <td><select value={state.status} onChange={(event) => updateState(selected, { status: event.target.value as ExamSetProblemState['status'] })}><option>未開始</option><option>開始</option><option>完了</option></select></td>
                    <td><input type="number" value={state.score} onChange={(event) => updateState(selected, { score: Number(event.target.value || 0) })} /></td>
                    <td><select value={state.rank} onChange={(event) => updateState(selected, { rank: event.target.value as ReviewRank })}><option value="">未選択</option><option value="A">A</option><option value="B">B</option><option value="C">C</option></select></td>
                    <td><button onClick={() => onOpenProblem(selected)}>問題へ移動</button></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="result-card"><strong>合計点：{totalScore}点</strong><span>{totalScore >= 70 ? '合格ライン到達' : '復習優先'}</span></div>
        <label>セットメモ<textarea value={memo} onChange={(event) => setMemo(event.target.value)} /></label>
        <button className="accent" onClick={finishAndSave}>セット演習を保存</button>

        <h3>セット履歴</h3>
        <div className="table-wrap compact-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>作成日時</th><th>セット名</th><th>合計点</th><th>判定</th><th>問題ID</th></tr></thead>
            <tbody>{examSets.slice(0, 10).map((set) => <tr key={set.id}><td>{set.createdAt}</td><td>{set.name}</td><td>{set.totalScore}</td><td>{set.passLineReached ? '合格ライン到達' : '復習優先'}</td><td>{set.selectedProblemIds.map((id) => problemCatalog.find((problem) => problem.id === id)?.displayId ?? id).join(' / ')}</td></tr>)}</tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
