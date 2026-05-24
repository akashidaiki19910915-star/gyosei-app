import { useState } from 'react';
import type { AnswerState, GradingStatus, MaterialPdf, MaterialPdfMapping, MissReason, ProblemDefinition, ReviewRank, TemplateDefinition } from '../types';
import { reviewDateForRank } from '../utils/dates';
import { AnswerTable } from './AnswerTable';
import { PdfViewerPanel } from './PdfViewerPanel';

interface Props {
  problem: ProblemDefinition;
  answer: AnswerState;
  template: TemplateDefinition;
  mapping?: MaterialPdfMapping;
  pdf?: MaterialPdf;
  onChange: (answer: AnswerState) => void;
  onSave: () => void;
  onApplyRowPoints: () => void;
  onSubmitSelfGrading: (input: { status: GradingStatus; score: number; maxScore: number; memo: string }) => void | Promise<void>;
  onBackPractice: () => void;
}

const gradingStatuses: GradingStatus[] = ['正解', '部分正解', '不正解', '迷いあり'];
const missReasons: MissReason[] = ['論点理解不足', '仕訳ミス', '借方貸方逆', '金額ミス', '集計ミス', '転記ミス', '表の入力位置ミス', '下書き不足', '時間不足', '解答欄形式の誤認', '問題文読み落とし', 'その他'];

function extractedAnswer(mapping?: MaterialPdfMapping): string {
  return [mapping?.answerText, mapping?.explanationText].filter(Boolean).join('\n\n') || '教材設定画面でPDFから解答・解説テキストを抽出すると、ここに端末内データとして表示されます。未設定の場合は、原本PDFまたは教材を見ながら自己採点してください。';
}

function nextRankFromStatus(status: GradingStatus): ReviewRank {
  if (status === '正解') return 'A';
  if (status === '部分正解' || status === '迷いあり') return 'B';
  return 'C';
}

export function FocusedGradingPanel({ problem, answer, template, mapping, pdf, onChange, onSave, onApplyRowPoints, onSubmitSelfGrading, onBackPractice }: Props) {
  const [status, setStatus] = useState<GradingStatus>('迷いあり');
  const [showOriginal, setShowOriginal] = useState(false);
  const [page, setPage] = useState(mapping?.answerPageStart ?? mapping?.answerPages?.[0] ?? mapping?.explanationPages?.[0] ?? 1);
  const explanation = extractedAnswer(mapping);

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
    <section className="focused-grading-panel">
      <section className="panel focused-question-card">
        <div className="focused-question-header">
          <div>
            <p className="eyebrow">採点する</p>
            <h2>{problem.displayId}：{problem.topic}</h2>
            <p>自分の回答を見ながら、端末内に抽出した解答・解説テキストで自己採点します。</p>
          </div>
          <div className="button-row focused-main-actions">
            <button className="secondary" onClick={onBackPractice}>答案入力へ戻る</button>
            {pdf && <button className="secondary" onClick={() => setShowOriginal((current) => !current)}>{showOriginal ? '原本を閉じる' : '原本を見る'}</button>}
            <button onClick={onSave}>一時保存</button>
            <button className="accent" onClick={() => { void onSubmitSelfGrading({ status, score: answer.score, maxScore: answer.maxScore, memo: answer.reviewMemo }); }}>保存して次へ進む</button>
          </div>
        </div>

        <div className="grading-text-grid">
          <div className="problem-text-box">
            <div className="problem-text-toolbar"><strong>解答・解説テキスト</strong><span>端末内PDF抽出データ</span></div>
            <div className="problem-text-content explanation-text-content">{explanation}</div>
          </div>
          <div className="grading-control-box">
            <h3>自己採点</h3>
            <div className="status-button-grid">
              {gradingStatuses.map((item) => <button key={item} className={status === item ? 'accent' : 'secondary'} onClick={() => applyStatus(item)}>{item}</button>)}
            </div>
            <div className="score-grid">
              <label>得点<input type="number" value={answer.score} onChange={(event) => onChange({ ...answer, score: Number(event.target.value || 0) })} /></label>
              <label>満点<input type="number" value={answer.maxScore} onChange={(event) => onChange({ ...answer, maxScore: Number(event.target.value || 0) })} /></label>
              <label>A/B/C判定<select value={answer.rank} onChange={(event) => setRank(event.target.value as ReviewRank)}><option value="">未選択</option><option value="A">A：自力で解けた</option><option value="B">B：手順に不安</option><option value="C">C：再復習</option></select></label>
              <label>次回復習日<input type="date" value={answer.nextReviewDate} onChange={(event) => onChange({ ...answer, nextReviewDate: event.target.value })} /></label>
            </div>
            <h3>ミス原因</h3>
            <div className="check-list compact-check-list">
              {missReasons.map((reason) => <label key={reason}><input type="checkbox" checked={answer.missReasons.includes(reason)} onChange={() => toggleReason(reason)} />{reason}</label>)}
            </div>
            <label>復習メモ<textarea value={answer.reviewMemo} onChange={(event) => onChange({ ...answer, reviewMemo: event.target.value })} placeholder="次回解く前に見る注意点" /></label>
          </div>
        </div>

        {showOriginal && <PdfViewerPanel pdf={pdf} title={pdf?.title ?? 'PDF原本'} label="解答・解説原本" page={page} pageStart={mapping?.answerPageStart ?? 1} pageEnd={mapping?.explanationPageEnd ?? pdf?.pageCount ?? 1} onPageChange={setPage} onClose={() => setShowOriginal(false)} />}
      </section>

      <section className="focused-answer-area">
        <div className="focused-answer-title"><div><p className="eyebrow">自分の回答</p><h2>{template.name}</h2></div><button className="secondary" onClick={onApplyRowPoints}>行別得点合計を反映</button></div>
        <AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} />
      </section>
    </section>
  );
}
