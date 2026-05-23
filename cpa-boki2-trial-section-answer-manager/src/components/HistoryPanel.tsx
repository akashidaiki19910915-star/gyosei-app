import type { HistoryEntry } from '../types';
import { formatDateTime, isDueTodayOrEarlier, isOverdue } from '../utils/dates';

type Filter = 'all' | 'today' | 'overdue' | 'c' | 'b';

interface Props {
  histories: HistoryEntry[];
  filter: Filter;
  onFilter: (filter: Filter) => void;
  onRestore: (entry: HistoryEntry) => void;
  onDelete: (id: string) => void;
  onClear: () => void;
}

export function HistoryPanel({ histories, filter, onFilter, onRestore, onDelete, onClear }: Props) {
  const filtered = histories.filter((history) => {
    if (filter === 'today') return isDueTodayOrEarlier(history.nextReviewDate);
    if (filter === 'overdue') return isOverdue(history.nextReviewDate);
    if (filter === 'c') return history.rank === 'C';
    if (filter === 'b') return history.rank === 'B';
    return true;
  });

  return (
    <section className="panel history-panel">
      <div className="section-heading">
        <h2>履歴一覧</h2>
        <div className="button-row">
          <button onClick={() => onFilter('all')}>全件</button>
          <button onClick={() => onFilter('today')}>今日復習</button>
          <button onClick={() => onFilter('overdue')}>期限超過</button>
          <button onClick={() => onFilter('c')}>C判定</button>
          <button onClick={() => onFilter('b')}>B判定</button>
          <button className="danger" onClick={onClear}>履歴全削除</button>
        </div>
      </div>
      <div className="table-wrap history-wrap">
        <table>
          <thead>
            <tr><th>保存日時</th><th>問題ID</th><th>論点</th><th>テンプレート</th><th>得点</th><th>判定</th><th>復習日</th><th>操作</th></tr>
          </thead>
          <tbody>
            {filtered.map((history) => (
              <tr key={history.id}>
                <td>{formatDateTime(history.savedAt)}</td>
                <td>{history.displayId}</td>
                <td>{history.topic}</td>
                <td>{history.templateName}</td>
                <td>{history.score}/{history.maxScore}</td>
                <td>{history.rank || '-'}</td>
                <td>{history.nextReviewDate || '-'}</td>
                <td className="button-cell"><button onClick={() => onRestore(history)}>復元</button><button className="danger" onClick={() => onDelete(history.id)}>削除</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}
