import { useMemo, useState } from 'react';
import type { HistoryEntry, MasteryFilter } from '../types';
import { formatDateTime } from '../utils/dates';
import { buildMasteryMap, filterMasteryMap } from '../utils/mastery';

interface Props {
  histories: HistoryEntry[];
  onOpenProblem: (problemId: string) => void;
  onExportCsv: () => void;
}

const filters: MasteryFilter[] = ['全体', '未着手', '危険問題', '期限超過', 'C判定', 'B判定', '合格水準', '商業簿記', '工業簿記', '第1問対策', '第2問対策', '第3問対策', '第4問対策', '第5問対策'];

function subjectLabel(value: string): string {
  return value === 'commercial' ? '商業簿記' : '工業簿記';
}

export function MasteryMapPanel({ histories, onOpenProblem, onExportCsv }: Props) {
  const [filter, setFilter] = useState<MasteryFilter>('危険問題');
  const items = useMemo(() => buildMasteryMap(histories), [histories]);
  const visible = useMemo(() => filterMasteryMap(items, filter), [filter, items]);
  const counts = {
    total: items.length,
    untouched: items.filter((item) => item.status === '未着手').length,
    dangerous: items.filter((item) => item.status === '危険問題' || item.status === '再復習対象' || item.status === '期限超過' || item.attempts === 0).length,
    mastered: items.filter((item) => item.status === '合格水準').length,
    overdue: items.filter((item) => item.status === '期限超過').length,
  };

  return (
    <details className="management-panel" open>
      <summary>合格到達マップ</summary>
      <div className="panel-body">
        <div className="metric-grid">
          <div><span>全問題ID</span><strong>{counts.total}</strong></div>
          <div><span>未着手</span><strong>{counts.untouched}</strong></div>
          <div><span>危険・再復習</span><strong>{counts.dangerous}</strong></div>
          <div><span>期限超過</span><strong>{counts.overdue}</strong></div>
          <div><span>合格水準</span><strong>{counts.mastered}</strong></div>
        </div>
        <div className="button-row">
          {filters.map((item) => <button key={item} className={filter === item ? 'accent' : 'secondary'} onClick={() => setFilter(item)}>{item}</button>)}
          <button onClick={onExportCsv}>合格到達マップCSV</button>
        </div>
        <div className="table-wrap compact-wrap">
          <table className="compact-table">
            <thead>
              <tr>
                <th>状態</th><th>問題ID</th><th>科目</th><th>大問</th><th>論点</th><th>最新得点</th><th>得点率</th><th>判定</th><th>回数</th><th>A/B/C</th><th>最終演習</th><th>次回復習</th><th>推奨アクション</th><th>開始</th>
              </tr>
            </thead>
            <tbody>
              {visible.map((item) => (
                <tr key={item.problem.id}>
                  <td><span className={`status-badge status-${item.status}`}>{item.status}</span></td>
                  <td>{item.problem.displayId}</td>
                  <td>{subjectLabel(item.problem.subject)}</td>
                  <td>{item.problem.sectionLabel}</td>
                  <td>{item.problem.topic}</td>
                  <td>{item.latestScore ?? '未'}/{item.maxScore ?? '-'}</td>
                  <td>{item.scoreRate === null ? '-' : `${item.scoreRate}%`}</td>
                  <td>{item.latestRank || '-'}</td>
                  <td>{item.attempts}</td>
                  <td>{item.aCount}/{item.bCount}/{item.cCount}</td>
                  <td>{formatDateTime(item.lastPracticedAt) || '-'}</td>
                  <td>{item.nextReviewDate || '-'}</td>
                  <td>{item.action}</td>
                  <td><button onClick={() => onOpenProblem(item.problem.id)}>開く</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
