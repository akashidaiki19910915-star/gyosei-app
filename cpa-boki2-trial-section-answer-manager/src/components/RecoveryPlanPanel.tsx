import { useMemo } from 'react';
import type { ExamSetRecord, HistoryEntry } from '../types';
import { buildRecoveryPlan } from '../utils/recovery';

interface Props {
  histories: HistoryEntry[];
  examSets: ExamSetRecord[];
  onExportCsv: () => void;
}

export function RecoveryPlanPanel({ histories, examSets, onExportCsv }: Props) {
  const plan = useMemo(() => buildRecoveryPlan(histories, examSets), [histories, examSets]);
  return (
    <details className="management-panel" open>
      <summary>70点突破リカバリー表</summary>
      <div className="panel-body">
        <div className="result-card">
          <strong>{plan.totalScore}/{plan.maxScore}点</strong>
          <span>判定元：{plan.source}</span>
          <span>70点まであと {plan.shortage70}点</span>
          <span>72点まであと {plan.shortage72}点</span>
          <span>80点まであと {plan.shortage80}点</span>
        </div>
        <div className="metric-grid">
          <div><span>最も失点が大きい大問</span><strong>{plan.biggestLossSection}</strong></div>
          <div><span>最も多いミス原因</span><strong>{plan.mostFrequentMissReason}</strong></div>
          <div><span>次回の重点対策</span><strong>{plan.nextFocus}</strong></div>
          <div><span>取り戻し候補</span><strong>{plan.recoveryCandidates.join(' / ') || 'データ不足'}</strong></div>
        </div>
        <div className="button-row"><button onClick={onExportCsv}>70点突破リカバリーCSV</button></div>
        <div className="table-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>大問</th><th>得点</th><th>満点</th><th>失点</th><th>対策優先度</th></tr></thead>
            <tbody>
              {plan.sectionRows.map((row) => (
                <tr key={row.sectionId}>
                  <td>{row.sectionLabel}</td><td>{row.score}</td><td>{row.maxScore}</td><td>{row.lostPoints}</td><td>{row.lostPoints >= 8 ? '最優先' : row.lostPoints >= 4 ? '優先' : '維持'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
