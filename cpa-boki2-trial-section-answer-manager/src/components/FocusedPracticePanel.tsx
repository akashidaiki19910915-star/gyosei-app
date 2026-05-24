import { useMemo, useState } from 'react';
import type { AnswerState, MaterialPdf, MaterialPdfMapping, ProblemDefinition, TemplateDefinition } from '../types';
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
  onGoGrading: () => void;
  onNextProblem: () => void;
  onApplyRowPoints: () => void;
}

function fallbackQuestion(problem: ProblemDefinition): string {
  return '教材を設定すると、PDFから抽出した問題文テキストがここに表示されます。未設定の場合は、購入済みPDFや紙教材を見ながら下の答案欄へ入力してください。';
}

function joinBlockText(mapping?: MaterialPdfMapping): string {
  const fromBlocks = mapping?.problemBlocks?.map((block) => block.questionText).filter(Boolean).join('\n\n');
  return fromBlocks || mapping?.problemText || '';
}

function pageLabel(pages?: number[]): string {
  if (!pages || pages.length === 0) return '未設定';
  return Array.from(new Set(pages)).sort((a, b) => a - b).join(', ');
}

export function FocusedPracticePanel({ problem, answer, template, mapping, pdf, onChange, onSave, onGoGrading, onNextProblem, onApplyRowPoints }: Props) {
  const [showOriginal, setShowOriginal] = useState(false);
  const [page, setPage] = useState(mapping?.problemPageStart ?? mapping?.sourcePages?.[0] ?? 1);
  const questionText = useMemo(() => joinBlockText(mapping) || fallbackQuestion(problem), [mapping, problem]);
  const subjectLabel = problem.subject === 'commercial' ? '商業簿記' : '工業簿記';
  const textReady = Boolean(joinBlockText(mapping));

  return (
    <section className="focused-practice-panel">
      <section className="panel focused-question-card">
        <div className="focused-question-header">
          <div>
            <p className="eyebrow">今すぐ解く</p>
            <h2>{problem.displayId}：{problem.topic}</h2>
            <div className="focused-question-meta">
              <span>{subjectLabel}</span>
              <span>{problem.sectionLabel}</span>
              <span>想定時間：{mapping?.estimatedMinutes ?? '未設定'}</span>
              <span>難易度：{mapping?.difficulty ?? '未設定'}</span>
              {mapping?.extractionConfidence !== undefined && <span>抽出信頼度：{mapping.extractionConfidence}%</span>}
            </div>
          </div>
          <div className="button-row focused-main-actions">
            <button onClick={onSave}>一時保存</button>
            <button className="accent" onClick={onGoGrading}>採点へ進む</button>
            <button className="secondary" onClick={onNextProblem}>次の問題へ</button>
          </div>
        </div>

        <div className="problem-text-box">
          <div className="problem-text-toolbar">
            <strong>問題文テキスト</strong>
            <span>抽出元ページ：{pageLabel(mapping?.sourcePages)}</span>
            {pdf && <button className="secondary" onClick={() => setShowOriginal((current) => !current)}>{showOriginal ? '原本を閉じる' : '原本を見る'}</button>}
          </div>
          {!textReady && <p className="warning-text">PDF抽出テキストが未設定です。教材設定画面でPDFから問題候補を抽出してください。</p>}
          <div className="problem-text-content">{questionText}</div>
        </div>

        {showOriginal && <PdfViewerPanel pdf={pdf} title={pdf?.title ?? 'PDF原本'} label="原本" page={page} pageStart={mapping?.problemPageStart ?? 1} pageEnd={mapping?.problemPageEnd ?? pdf?.pageCount ?? 1} onPageChange={setPage} onClose={() => setShowOriginal(false)} />}
      </section>

      <section className="focused-answer-area">
        <div className="focused-answer-title">
          <div>
            <p className="eyebrow">答案入力</p>
            <h2>{template.name}</h2>
          </div>
          <button className="secondary" onClick={onApplyRowPoints}>行別得点合計を反映</button>
        </div>
        <AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} />
      </section>
    </section>
  );
}
