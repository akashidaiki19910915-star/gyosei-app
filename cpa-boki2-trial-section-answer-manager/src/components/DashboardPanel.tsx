import type { HistoryEntry } from '../types';
import { buildDashboard } from '../utils/dashboard';

interface Props {
  histories: HistoryEntry[];
  answerCount: number;
  onExportCsv: () => void;
}

function pct(value: number): string {
  return `${Math.round(value * 100)}%`;
}

export function DashboardPanel({ histories, answerCount, onExportCsv }: Props) {
  const dashboard = buildDashboard(histories, answerCount);
  const { summary } = dashboard;

  return (
    <details className="panel management-panel">
      <summary>成績ダッシュボード</summary>
      <div className="panel-body">
        <div className="button-row"><button onClick={onExportCsv}>成績ダッシュボードCSV</button></div>
        <div className="metric-grid">
          <div><span>総演習回数</span><strong>{summary.totalAttempts}</strong></div>
          <div><span>履歴件数</span><strong>{summary.historyCount}</strong></div>
          <div><span>保存済み答案</span><strong>{summary.answerCount}</strong></div>
          <div><span>平均点</span><strong>{summary.averageScore.toFixed(1)}</strong></div>
          <div><span>平均得点率</span><strong>{pct(summary.averageRate)}</strong></div>
          <div><span>70%以上相当</span><strong>{summary.passLikeCount}</strong></div>
          <div><span>A判定</span><strong>{summary.aCount}</strong></div>
          <div><span>B判定</span><strong>{summary.bCount}</strong></div>
          <div><span>C判定</span><strong>{summary.cCount}</strong></div>
          <div><span>未着手</span><strong>{summary.untouchedCount}</strong></div>
          <div><span>期限超過</span><strong>{summary.overdueCount}</strong></div>
          <div><span>今日復習</span><strong>{summary.todayCount}</strong></div>
          <div><span>直近7日演習</span><strong>{summary.sevenDayAttempts}</strong></div>
          <div><span>直近30日演習</span><strong>{summary.thirtyDayAttempts}</strong></div>
          <div><span>直近7日C</span><strong>{summary.sevenDayC}</strong></div>
          <div><span>直近30日C</span><strong>{summary.thirtyDayC}</strong></div>
        </div>

        <h3>問題ID別成績</h3>
        <div className="table-wrap compact-wrap">
          <table className="compact-table">
            <thead><tr><th>問題ID</th><th>論点</th><th>回数</th><th>最新</th><th>最高</th><th>最低</th><th>判定</th><th>ミス原因</th><th>復習日</th><th>改善</th></tr></thead>
            <tbody>
              {dashboard.problemStats.map((row) => (
                <tr key={row.problem.id}>
                  <td>{row.problem.displayId}</td><td>{row.problem.topic}</td><td>{row.attempts}</td><td>{row.latestScore ?? '-'}</td><td>{row.highestScore ?? '-'}</td><td>{row.lowestScore ?? '-'}</td><td>{row.latestRank || '-'}</td><td>{row.latestMissReasons.join(' / ') || '-'}</td><td>{row.latestReviewDate || '-'}</td><td>{row.trend}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <h3>大問別成績</h3>
        <div className="table-wrap compact-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>大問</th><th>演習回数</th><th>平均点</th><th>C判定</th><th>よくあるミス原因</th></tr></thead>
            <tbody>{dashboard.sectionStats.map((row) => <tr key={row.sectionId}><td>{row.sectionLabel}</td><td>{row.attempts}</td><td>{row.averageScore.toFixed(1)}</td><td>{row.cCount}</td><td>{row.commonReason}</td></tr>)}</tbody>
          </table>
        </div>

        <h3>ミス原因ランキング</h3>
        <div className="table-wrap compact-wrap short-wrap">
          <table className="compact-table">
            <thead><tr><th>ミス原因</th><th>件数</th><th>該当問題ID</th><th>最新発生日</th></tr></thead>
            <tbody>{dashboard.missReasonRanking.map((row) => <tr key={row.reason}><td>{row.reason}</td><td>{row.count}</td><td>{row.problemIds.join(' / ')}</td><td>{row.latestAt}</td></tr>)}</tbody>
          </table>
        </div>
      </div>
    </details>
  );
}
