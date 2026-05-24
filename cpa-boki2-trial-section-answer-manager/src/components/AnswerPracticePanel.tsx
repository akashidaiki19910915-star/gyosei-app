import { useState } from 'react';
import type { AnswerState, MaterialPdf, MaterialPdfMapping, ProblemDefinition, QuestionCard, TemplateDefinition, TemplateId } from '../types';
import { AnswerTable } from './AnswerTable';
import { JournalEntryCardInput } from './JournalEntryCardInput';
import { PdfViewerPanel } from './PdfViewerPanel';
import { ProblemSelector } from './ProblemSelector';

interface Props {
  problem: ProblemDefinition;
  questionCard: QuestionCard;
  mapping?: MaterialPdfMapping;
  pdf?: MaterialPdf;
  answer: AnswerState;
  template: TemplateDefinition;
  progressIndex: number;
  totalToday: number;
  onSelectProblem: (problemId: string) => void;
  onSelectTemplate: (templateId: TemplateId) => void;
  onChange: (answer: AnswerState) => void;
  onSave: () => void;
  onGoGrading: () => void;
  onNextProblem: () => void;
  onApplyRowPoints: () => void;
}

function hasUsableQuestionText(card: QuestionCard): boolean {
  return Boolean(card.questionText && !card.questionText.includes('問題文未登録') && card.questionText.trim().length > 20);
}

function pageStart(card: QuestionCard, mapping?: MaterialPdfMapping): number {
  return card.sourcePageStart || mapping?.problemPageStart || mapping?.sourcePages?.[0] || 1;
}

function pageEnd(card: QuestionCard, mapping?: MaterialPdfMapping, pdf?: MaterialPdf): number {
  return card.sourcePageEnd || mapping?.problemPageEnd || pdf?.pageCount || pageStart(card, mapping);
}

export function AnswerPracticePanel({ problem, questionCard, mapping, pdf, answer, template, progressIndex, totalToday, onSelectProblem, onSelectTemplate, onChange, onSave, onGoGrading, onNextProblem, onApplyRowPoints }: Props) {
  const [showOriginal, setShowOriginal] = useState(false);
  const [page, setPage] = useState(pageStart(questionCard, mapping));
  const showQuestionText = hasUsableQuestionText(questionCard) && !questionCard.needsReview;

  return (
    <section className="answer-practice-shell">
      <section className="panel practice-hero-card">
        <div>
          <p className="eyebrow">今すぐ解く</p>
          <h2>{problem.displayId}：{problem.topic}</h2>
          <p className="practice-notice">問題を開く → 答案を入力する → 採点する → 復習日に回す、の順で進めます。</p>
        </div>
        <div className="practice-progress-box">
          <strong>{Math.max(1, progressIndex)} / {Math.max(1, totalToday)}</strong>
          <span>現在の問題番号</span>
        </div>
      </section>

      <div className="answer-practice-layout">
        <ProblemSelector selectedProblem={problem} selectedTemplateId={answer.templateId} onSelectProblem={onSelectProblem} onSelectTemplate={onSelectTemplate} />

        <section className="answer-workspace">
          <section className="panel answer-summary-card">
            <div className="answer-summary-grid">
              <div><span>科目</span><strong>{problem.subject}</strong></div>
              <div><span>大問対策</span><strong>{problem.sectionLabel}</strong></div>
              <div><span>問題ID</span><strong>{problem.displayId}</strong></div>
              <div><span>論点名</span><strong>{problem.topic}</strong></div>
              <div><span>想定時間</span><strong>{questionCard.estimatedMinutes ? `${questionCard.estimatedMinutes}分` : '教材側で確認'}</strong></div>
              <div><span>テンプレート</span><strong>{template.name}</strong></div>
            </div>
            <div className="button-row answer-primary-actions">
              <button onClick={onSave}>一時保存</button>
              <button className="accent" onClick={onGoGrading}>採点へ進む</button>
              <button className="secondary" onClick={onNextProblem}>次の問題へ</button>
            </div>
          </section>

          <section className="panel question-reference-card">
            <div className="problem-text-toolbar">
              <strong>問題文</strong>
              {pdf && <button className="secondary" onClick={() => setShowOriginal((current) => !current)}>{showOriginal ? '原本を閉じる' : '原本を見る'}</button>}
            </div>
            {showQuestionText ? (
              <div className="problem-text-content readable-question-text">{questionCard.questionText}</div>
            ) : (
              <p className="practice-notice">問題文はお手元の教材・PDFで確認してください。この画面では答案入力と復習管理を行います。</p>
            )}
            {showOriginal && <PdfViewerPanel pdf={pdf} title={pdf?.title ?? 'PDF原本'} label="原本を見る" page={page} pageStart={pageStart(questionCard, mapping)} pageEnd={pageEnd(questionCard, mapping, pdf)} onPageChange={setPage} onClose={() => setShowOriginal(false)} />}
          </section>

          <section className="focused-answer-area answer-input-main answer-only-main">
            <div className="focused-answer-title">
              <div>
                <p className="eyebrow">答案入力</p>
                <h2>{template.name}</h2>
              </div>
              <button className="secondary" onClick={onApplyRowPoints}>行別得点合計を問題得点へ反映</button>
            </div>
            {answer.templateId === 'journal' && <JournalEntryCardInput answer={answer} template={template} onChange={onChange} />}
            <div className="pc-answer-only"><AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} /></div>
          </section>

          <section className="panel draft-memo-panel">
            <label>下書きメモ・計算過程
              <textarea value={answer.draftMemo} onChange={(event) => onChange({ ...answer, draftMemo: event.target.value })} placeholder="本・PDFを見ながら、仕訳メモ、計算過程、次に確認する点だけを残します。" />
            </label>
          </section>
        </section>
      </div>
    </section>
  );
}
