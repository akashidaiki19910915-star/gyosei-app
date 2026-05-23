import type { AnswerState, MissReason, ReviewRank } from '../types';
import { reviewDateForRank } from '../utils/dates';

const missReasons: MissReason[] = [
  '論点理解不足',
  '仕訳ミス',
  '借方貸方逆',
  '金額ミス',
  '集計ミス',
  '転記ミス',
  '表の入力位置ミス',
  '下書き不足',
  '時間不足',
  '解答欄形式の誤認',
  '問題文読み落とし',
  'その他',
];

interface Props {
  answer: AnswerState;
  onChange: (answer: AnswerState) => void;
  onSave: () => void;
  onConfirmScoring: () => void;
}

export function ScoringPanel({ answer, onChange, onSave, onConfirmScoring }: Props) {
  const setRank = (rank: ReviewRank) => onChange({ ...answer, rank, nextReviewDate: reviewDateForRank(rank) });
  const toggleReason = (reason: MissReason) => {
    const exists = answer.missReasons.includes(reason);
    onChange({ ...answer, missReasons: exists ? answer.missReasons.filter((item) => item !== reason) : [...answer.missReasons, reason] });
  };

  return (
    <aside className="panel scoring-panel">
      <h2>自己採点・復習管理</h2>
      <div className="score-grid">
        <label>問題得点<input type="number" value={answer.score} onChange={(event) => onChange({ ...answer, score: Number(event.target.value || 0) })} /></label>
        <label>満点<input type="number" value={answer.maxScore} onChange={(event) => onChange({ ...answer, maxScore: Number(event.target.value || 0) })} /></label>
        <label>行別得点合計<input type="number" value={answer.rowPointsTotal} readOnly /></label>
        <label>A/B/C判定
          <select value={answer.rank} onChange={(event) => setRank(event.target.value as ReviewRank)}>
            <option value="">未選択</option>
            <option value="A">A：自力で解けた・説明できる</option>
            <option value="B">B：惜しいが手順に不安</option>
            <option value="C">C：再現できない</option>
          </select>
        </label>
        <label>次回復習日<input type="date" value={answer.nextReviewDate} onChange={(event) => onChange({ ...answer, nextReviewDate: event.target.value })} /></label>
      </div>

      <h3>ミス原因</h3>
      <div className="check-list">
        {missReasons.map((reason) => (
          <label key={reason}><input type="checkbox" checked={answer.missReasons.includes(reason)} onChange={() => toggleReason(reason)} />{reason}</label>
        ))}
      </div>

      <label>下書きメモ<textarea value={answer.draftMemo} onChange={(event) => onChange({ ...answer, draftMemo: event.target.value })} /></label>
      <label>復習メモ<textarea value={answer.reviewMemo} onChange={(event) => onChange({ ...answer, reviewMemo: event.target.value })} /></label>

      <div className="score-mode">採点モード：{answer.scoringStarted ? '開始済み' : '未開始'}</div>
      <div className="button-grid">
        <button className="secondary" onClick={() => onChange({ ...answer, scoringStarted: true })}>採点モード開始</button>
        <button onClick={onSave}>手動保存</button>
        <button className="accent" onClick={onConfirmScoring}>採点確定して履歴追加</button>
      </div>
    </aside>
  );
}
