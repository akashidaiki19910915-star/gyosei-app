import type { StudyQueueItem } from '../types';
import { queueCounts } from '../utils/studyQueue';

interface Props {
  items: StudyQueueItem[];
  onStart: (problemId: string) => void;
  onRestoreLatest?: (item: StudyQueueItem) => void;
  onExportCsv: () => void;
}

export function StudyQueuePanel({ items, onStart, onRestoreLatest, onExportCsv }: Props) {
  const counts = queueCounts(items);
  const top = items[0];

  return (
    <details className="panel management-panel" open>
      <summary>今日の学習キュー：{counts.total}件 / 期限超過 {counts.overdue}件 / C判定 {counts.c}件 / 未着手 {counts.untouched}件</summary>
      <div className="panel-body">
        {top ? (
          <div className="priority-card">
            <div className="badge-row">{top.statusLabels.map((label) => <span key={label} className="status-badge">{label}</span>)}</div>
            <h3>今日の最優先：{top.problem.displayId} {top.problem.topic}</h3>
            <p>{top.reason}</p>
            <button onClick={() => onStart(top.problem.id)}>この問題を開始</button>
          </div>
        ) : <p className="empty">今日の学習キューはありません。</p>}

        <div className="button-row"><button onClick={onExportCsv}>今日の学習キューCSV</button></div>
        <div className="table-wrap compact-wrap">
          <table className="compact-table">
            <thead>
              <tr><th>優先</th><th>状態</th><th>問題ID</th><th>科目</th><th>大問</th><th>論点</th><th>前回得点</th><th>判定</th><th>復習日</th><th>ミス原因</th><th>メモ</th><th>操作</th></tr>
            </thead>
            <tbody>
              {items.slice(0, 30).map((item) => (
                <tr key={item.problem.id}>
                  <td>{item.priority}</td>
                  <td>{item.statusLabels.join(' / ')}</td>
                  <td>{item.problem.displayId}</td>
                  <td>{item.problem.subject === 'commercial' ? '商業簿記' : '工業簿記'}</td>
                  <td>{item.problem.sectionLabel}</td>
                  <td>{item.problem.topic}</td>
                  <td>{item.previousScore ?? '-'}/{item.maxScore ?? '-'}</td>
                  <td>{item.rank || '-'}</td>
                  <td>{item.nextReviewDate || '-'}</td>
                  <td>{item.missReasons.join(' / ') || '-'}</td>
                  <td>{item.memoSummary || '-'}</td>
                  <td className="button-cell"><button onClick={() => onStart(item.problem.id)}>開始</button>{item.latestHistory && onRestoreLatest ? <button className="secondary" onClick={() => onRestoreLatest(item)}>最新履歴</button> : null}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
