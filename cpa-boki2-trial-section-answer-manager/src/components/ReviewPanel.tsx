import type { HistoryEntry } from '../types';
import { formatDateTime, isDueTodayOrEarlier, isOverdue } from '../utils/dates';

interface Props {
  histories: HistoryEntry[];
}

function RowList({ title, rows }: { title: string; rows: HistoryEntry[] }) {
  return (
    <section className="review-box">
      <h3>{title}</h3>
      {rows.length === 0 ? <p className="empty">該当なし</p> : rows.slice(0, 12).map((history) => (
        <div key={history.id} className="review-row">
          <strong>{history.displayId}</strong> {history.topic}<br />
          <span>{history.rank || '未判定'} / 次回: {history.nextReviewDate || '-'} / {formatDateTime(history.savedAt)}</span>
        </div>
      ))}
    </section>
  );
}

export function ReviewPanel({ histories }: Props) {
  const today = histories.filter((history) => isDueTodayOrEarlier(history.nextReviewDate));
  const overdue = histories.filter((history) => isOverdue(history.nextReviewDate));
  const cRows = histories.filter((history) => history.rank === 'C');
  const bRows = histories.filter((history) => history.rank === 'B');
  return (
    <aside className="panel review-panel">
      <h2>復習対象</h2>
      <RowList title="今日の復習対象" rows={today} />
      <RowList title="期限超過" rows={overdue} />
      <RowList title="C判定一覧" rows={cRows} />
      <RowList title="B判定一覧" rows={bRows} />
    </aside>
  );
}
