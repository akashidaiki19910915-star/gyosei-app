import { useMemo, useState } from 'react';
import type { AnswerState, MaterialPdf, MaterialPdfMapping, ProblemDefinition, QuestionCard, TemplateDefinition } from '../types';
import { AnswerTable } from './AnswerTable';
import { JournalEntryCardInput } from './JournalEntryCardInput';
import { PdfViewerPanel } from './PdfViewerPanel';

interface Props {
  problem: ProblemDefinition;
  questionCard: QuestionCard;
  answer: AnswerState;
  template: TemplateDefinition;
  mapping?: MaterialPdfMapping;
  pdf?: MaterialPdf;
  totalToday: number;
  progressIndex: number;
  onChange: (answer: AnswerState) => void;
  onSave: () => void;
  onGoGrading: () => void;
  onNextProblem: () => void;
  onApplyRowPoints: () => void;
}

function pageLabel(card: QuestionCard, mapping?: MaterialPdfMapping): string {
  if (card.sourcePageStart && card.sourcePageEnd) return card.sourcePageStart === card.sourcePageEnd ? String(card.sourcePageStart) : `${card.sourcePageStart}-${card.sourcePageEnd}`;
  const pages = mapping?.sourcePages;
  if (!pages || pages.length === 0) return '未設定';
  return Array.from(new Set(pages)).sort((a, b) => a - b).join(', ');
}

function hasBlankRequiredJournalCells(answer: AnswerState): boolean {
  if (answer.templateId !== 'journal') return answer.rows.every((row) => row.cells.every((cell, index) => index === 0 || cell.trim() === ''));
  const importantIndexes = answer.columns.map((column, index) => ({ column, index })).filter(({ column }) => column.includes('借方科目') || column.includes('借方金額') || column.includes('貸方科目') || column.includes('貸方金額')).map(({ index }) => index);
  return answer.rows.some((row) => importantIndexes.some((index) => (row.cells[index] ?? '').trim() === ''));
}

export function FocusedPracticePanel({ problem, questionCard, answer, template, mapping, pdf, totalToday, progressIndex, onChange, onSave, onGoGrading, onNextProblem, onApplyRowPoints }: Props) {
  const [showOriginal, setShowOriginal] = useState(false);
  const [page, setPage] = useState(questionCard.sourcePageStart || mapping?.problemPageStart || mapping?.sourcePages?.[0] || 1);
  const textReady = questionCard.questionText && !questionCard.questionText.includes('問題文未登録');
  const bottomLabel = useMemo(() => hasBlankRequiredJournalCells(answer) ? '未入力があります。採点しますか？' : '採点する', [answer]);
  const pageStart = questionCard.sourcePageStart || mapping?.problemPageStart || 1;
  const pageEnd = questionCard.sourcePageEnd || mapping?.problemPageEnd || pdf?.pageCount || 1;

  return (
    <section className="question-card-learning-flow">
      <div className="learning-top-bar">
        <button className="secondary" onClick={onNextProblem}>戻る</button>
        <div>
          <p className="eyebrow">今すぐ解く</p>
          <strong>{questionCard.topic}</strong>
        </div>
        <span>{Math.max(1, progressIndex)} / {Math.max(1, totalToday)}</span>
      </div>

      <section className="panel focused-question-card question-card-main">
        <div className="focused-question-header">
          <div>
            <p className="question-label">問題</p>
            <h2>{questionCard.title || `${problem.displayId}：${problem.topic}`}</h2>
            <div className="focused-question-meta subtle-meta">
              <span>{questionCard.subject}</span>
              <span>{questionCard.section}</span>
              <span>想定{questionCard.estimatedMinutes}分</span>
              <span>難易度：{questionCard.difficulty}</span>
              <span>信頼度：{questionCard.extractionConfidence}%</span>
              {questionCard.needsReview && <span className="needs-review-pill">要確認</span>}
            </div>
          </div>
          <div className="button-row focused-main-actions desktop-only-actions">
            <button onClick={onSave}>一時保存</button>
            <button className="accent" onClick={onGoGrading}>採点へ進む</button>
            <button className="secondary" onClick={onNextProblem}>次の問題へ</button>
          </div>
        </div>

        <div className="problem-text-box app-like-problem-box">
          <div className="problem-text-toolbar">
            <strong>問題文</strong>
            <span>原本P{pageLabel(questionCard, mapping)}</span>
            {pdf && <button className="secondary" onClick={() => setShowOriginal((current) => !current)}>{showOriginal ? '原本を閉じる' : '原本を見る'}</button>}
          </div>
          {!textReady && <p className="warning-text">問題文が未登録または要確認です。教材を設定する画面で問題カードを確認してください。</p>}
          <div className="problem-text-content app-like-problem-text">{questionCard.questionText}</div>
          {questionCard.choices.length > 0 && <div className="choices-box"><strong>選択肢</strong>{questionCard.choices.map((choice) => <span key={choice}>{choice}</span>)}</div>}
        </div>

        {showOriginal && <PdfViewerPanel pdf={pdf} title={pdf?.title ?? 'PDF原本'} label="原本確認" page={page} pageStart={pageStart} pageEnd={pageEnd} onPageChange={setPage} onClose={() => setShowOriginal(false)} />}
      </section>

      <section className="focused-answer-area answer-input-main">
        <div className="focused-answer-title">
          <div>
            <p className="eyebrow">答案入力</p>
            <h2>{questionCard.answerTemplateType}</h2>
          </div>
          <button className="secondary pc-answer-only" onClick={onApplyRowPoints}>行別得点合計を反映</button>
        </div>
        {questionCard.answerInputType === 'journalEntry' ? <JournalEntryCardInput answer={answer} template={template} onChange={onChange} /> : null}
        <div className="pc-answer-only"><AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} /></div>
      </section>

      <div className="mobile-bottom-action"><button className="accent" onClick={onGoGrading}>{bottomLabel}</button></div>
    </section>
  );
}
