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

interface StructuredQuestionText {
  instruction: string;
  body: string;
  notes: string[];
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

function formatMoneyInText(text: string): string {
  return text.replace(/\b\d{4,}\b/g, (value) => Number(value).toLocaleString('ja-JP'));
}

function structureQuestionText(text: string): StructuredQuestionText {
  const normalized = formatMoneyInText(text).replace(/\r\n?/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  const lines = normalized.split('\n').map((line) => line.trim()).filter(Boolean);
  const notes = lines.filter((line) => /^※|^注|^なお/.test(line));
  const mainLines = lines.filter((line) => !notes.includes(line));
  const joined = mainLines.join('\n\n');
  const firstSentenceMatch = joined.match(/^(.{8,120}?(?:しなさい。|答えなさい。|選びなさい。|選択しなさい。|記入しなさい。))/);
  if (firstSentenceMatch) {
    return {
      instruction: firstSentenceMatch[1],
      body: joined.slice(firstSentenceMatch[1].length).trim(),
      notes,
    };
  }
  return {
    instruction: '次の問題を解いてください。',
    body: joined,
    notes,
  };
}

export function FocusedPracticePanel({ problem, questionCard, answer, template, mapping, pdf, totalToday, progressIndex, onChange, onSave, onGoGrading, onNextProblem, onApplyRowPoints }: Props) {
  const [showOriginal, setShowOriginal] = useState(false);
  const [page, setPage] = useState(questionCard.sourcePageStart || mapping?.problemPageStart || mapping?.sourcePages?.[0] || 1);
  const textReady = questionCard.questionText && !questionCard.questionText.includes('問題文未登録');
  const bottomLabel = useMemo(() => hasBlankRequiredJournalCells(answer) ? '未入力があります。採点しますか？' : '解答する', [answer]);
  const pageStart = questionCard.sourcePageStart || mapping?.problemPageStart || 1;
  const pageEnd = questionCard.sourcePageEnd || mapping?.problemPageEnd || pdf?.pageCount || 1;
  const structuredText = useMemo(() => structureQuestionText(questionCard.questionText), [questionCard.questionText]);

  return (
    <section className="question-card-learning-flow study-app-shell">
      <div className="learning-top-bar study-app-top-bar">
        <button className="secondary" onClick={onNextProblem}>問題一覧へ</button>
        <div className="study-position-block">
          <p className="eyebrow">今すぐ解く</p>
          <strong>{questionCard.subject} / {questionCard.section}</strong>
          <small>{questionCard.topic}</small>
        </div>
        <div className="study-progress-block" aria-label="進捗">
          <span>{Math.max(1, progressIndex)} / {Math.max(1, totalToday)}</span>
          <div className="study-progress-track"><i style={{ width: `${Math.min(100, Math.max(4, (Math.max(1, progressIndex) / Math.max(1, totalToday)) * 100))}%` }} /></div>
        </div>
      </div>

      <section className="panel focused-question-card question-card-main study-question-card">
        <div className="focused-question-header">
          <div>
            <p className="question-label">問題</p>
            <h2>{questionCard.questionNumber}：{questionCard.topic || problem.topic}</h2>
            <div className="focused-question-meta subtle-meta">
              <span>{questionCard.subject}</span>
              <span>{questionCard.section}</span>
              <span>想定{questionCard.estimatedMinutes}分</span>
              <span>難易度：{questionCard.difficulty}</span>
              <span>信頼度：{questionCard.extractionConfidence}%</span>
              {questionCard.needsReview && <span className="needs-review-pill">原本確認が必要</span>}
            </div>
          </div>
          <div className="button-row focused-main-actions desktop-only-actions">
            <button onClick={onSave}>一時保存</button>
            <button className="accent" onClick={onGoGrading}>解答する</button>
            <button className="secondary" onClick={onNextProblem}>次の問題へ</button>
          </div>
        </div>

        <div className="problem-text-box app-like-problem-box structured-question-box">
          <div className="problem-text-toolbar">
            <strong>問題文</strong>
            <span>原本P{pageLabel(questionCard, mapping)}</span>
            {pdf && <button className="secondary" onClick={() => setShowOriginal((current) => !current)}>{showOriginal ? '原本を閉じる' : '原本を見る'}</button>}
          </div>
          {!textReady && <p className="warning-text">問題文が未登録または要確認です。教材を設定する画面で問題カードを確認してください。</p>}
          <div className="structured-question-content">
            <section className="question-part question-instruction"><h3>問題指示</h3><p>{structuredText.instruction}</p></section>
            <section className="question-part question-body"><h3>取引・条件・資料</h3><div className="problem-text-content app-like-problem-text">{structuredText.body || questionCard.questionText}</div></section>
            {questionCard.choices.length > 0 && <section className="question-part choices-box"><h3>選択肢</h3><div>{questionCard.choices.map((choice) => <span key={choice}>{choice}</span>)}</div></section>}
            {structuredText.notes.length > 0 && <section className="question-part note-box"><h3>注意書き</h3>{structuredText.notes.map((note) => <p key={note}>{note}</p>)}</section>}
          </div>
        </div>

        {showOriginal && <PdfViewerPanel pdf={pdf} title={pdf?.title ?? 'PDF原本'} label="原本確認" page={page} pageStart={pageStart} pageEnd={pageEnd} onPageChange={setPage} onClose={() => setShowOriginal(false)} />}
      </section>

      <section className="focused-answer-area answer-input-main study-answer-card">
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

      <div className="mobile-bottom-action three-button-footer">
        <button className="secondary" onClick={onSave}>メモ/保存</button>
        <button className="accent" onClick={onGoGrading}>{bottomLabel}</button>
        <button className="secondary" onClick={onNextProblem}>次へ</button>
      </div>
    </section>
  );
}
