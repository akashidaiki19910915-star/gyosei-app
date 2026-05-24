import { useState } from 'react';
import type { AnswerState, GradingStatus, MissReason, ProblemDefinition, ReviewRank, TemplateDefinition } from '../types';
import { reviewDateForRank } from '../utils/dates';
import { AnswerTable } from './AnswerTable';

interface Props {
  problem: ProblemDefinition;
  answer: AnswerState;
  template: TemplateDefinition;
  onChange: (answer: AnswerState) => void;
  onSave: () => void;
  onApplyRowPoints: () => void;
  onSubmitSelfGrading: (input: { status: GradingStatus; score: number; maxScore: number; memo: string }) => void | Promise<void>;
  onBackPractice: () => void;
}

const gradingStatuses: GradingStatus[] = ['正解', '部分正解', '不正解', '迷いあり'];
const missReasons: { value: MissReason; label: string }[] = [
  { value: '論点理解不足', label: '論点理解不足' },
  { value: '仕訳ミス', label: '仕訳ミス' },
  { value: '借方貸方逆', label: '借方貸方逆' },
  { value: '金額ミス', label: '金額ミス' },
  { value: '集計ミス', label: '集計ミス' },
  { value: '転記ミス', label: '転記ミス' },
  { value: '表の入力位置ミス', label: '表の入力位置ミス' },
  { value: '下書き不足', label: '下書き不足' },
  { value: '時間不足', label: '時間不足' },
  { value: '解答欄形式の誤認', label: '解答形式の誤認' },
  { value: '問題文読み落とし', label: '問題文読み落とし' },
  { value: 'その他', label: 'その他' },
];

function nextRankFromStatus(status: GradingStatus): ReviewRank {
  if (status === '正解') return 'A';
  if (status === '部分正解' || status === '迷いあり') return 'B';
  return 'C';
}

export function FocusedGradingPanel({ problem, answer, template, onChange, onSave, onApplyRowPoints, onSubmitSelfGrading, onBackPractice }: Props) {
  const [status, setStatus] = useState<GradingStatus>('迷いあり');

  const setRank = (rank: ReviewRank) => onChange({ ...answer, rank, nextReviewDate: reviewDateForRank(rank) });
  const toggleReason = (reason: MissReason) => {
    const exists = answer.missReasons.includes(reason);
    onChange({ ...answer, missReasons: exists ? answer.missReasons.filter((item) => item !== reason) : [...answer.missReasons, reason] });
  };
  const applyStatus = (nextStatus: GradingStatus) => {
    setStatus(nextStatus);
    const nextRank = nextRankFromStatus(nextStatus);
    setRank(nextRank);
  };

  return (
    <section className="focused-grading-panel self-grading-only">
      <section className="panel focused-question-card">
        <div className="focused-question-header">
          <div>
            <p className="eyebrow">採点する</p>
            <h2>{problem.displayId}：{problem.topic}</h2>
            <p>解答・解説はお手元の教材・PDFで確認し、この画面では自己採点、ミス原因、次回復習日だけを保存します。</p>
          </div>
          <div className="button-row focused-main-actions">
            <button className="secondary" onClick={onBackPractice}>答案入力へ戻る</button>
            <button onClick={onSave}>一時保存</button>
            <button className="accent" onClick={() => { void onSubmitSelfGrading({ status, score: answer.score, maxScore: answer.maxScore, memo: answer.reviewMemo }); }}>保存して次へ</button>
          </div>
        </div>

        <div className="grading-text-grid simplified-grading-grid">
          <div className="grading-control-box self-score-box">
            <h3>自己採点</h3>
            <p className="notice-small">自動採点は行いません。教材の解答・解説を見ながら、得点と判定を入力してください。</p>
            <div className="status-button-grid">
              {gradingStatuses.map((item) => <button key={item} className={status === item ? 'accent' : 'secondary'} onClick={() => applyStatus(item)}>{item}</button>)}
            </div>
            <div className="score-grid">
              <label>問題得点<input type="number" value={answer.score} onChange={(event) => onChange({ ...answer, score: Number(event.target.value || 0) })} /></label>
              <label>満点<input type="number" value={answer.maxScore} onChange={(event) => onChange({ ...answer, maxScore: Number(event.target.value || 0) })} /></label>
              <label>A/B/C判定<select value={answer.rank} onChange={(event) => setRank(event.target.value as ReviewRank)}><option value="">未選択</option><option value="A">A：自力で解けた</option><option value="B">B：手順に不安</option><option value="C">C：再復習</option></select></label>
              <label>次回復習日<input type="date" value={answer.nextReviewDate} onChange={(event) => onChange({ ...answer, nextReviewDate: event.target.value })} /></label>
            </div>
            <h3>ミス原因</h3>
            <div className="check-list compact-check-list">
              {missReasons.map((reason) => <label key={reason.value}><input type="checkbox" checked={answer.missReasons.includes(reason.value)} onChange={() => toggleReason(reason.value)} />{reason.label}</label>)}
            </div>
            <label>復習メモ<textarea value={answer.reviewMemo} onChange={(event) => onChange({ ...answer, reviewMemo: event.target.value })} placeholder="次回解く前に見る注意点" /></label>
          </div>
        </div>
      </section>

      <section className="focused-answer-area">
        <div className="focused-answer-title"><div><p className="eyebrow">自分の答案</p><h2>{template.name}</h2></div><button className="secondary" onClick={onApplyRowPoints}>行別得点合計を反映</button></div>
        <AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} />
      </section>
    </section>
  );
}
