import type { AnswerState, ProblemDefinition, TemplateDefinition, TemplateId } from '../types';
import { AnswerTable } from './AnswerTable';
import { ProblemSelector } from './ProblemSelector';

interface Props {
  problem: ProblemDefinition;
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

export function AnswerPracticePanel({ problem, answer, template, progressIndex, totalToday, onSelectProblem, onSelectTemplate, onChange, onSave, onGoGrading, onNextProblem, onApplyRowPoints }: Props) {
  return (
    <section className="answer-practice-shell">
      <section className="panel practice-hero-card">
        <div>
          <p className="eyebrow">解く</p>
          <h2>{problem.displayId}：{problem.topic}</h2>
          <p className="practice-notice">問題文はお手元の教材・PDFで確認してください。この画面では答案入力と復習管理だけを行います。</p>
        </div>
        <div className="practice-progress-box">
          <strong>{Math.max(1, progressIndex)} / {Math.max(1, totalToday)}</strong>
          <span>今日の学習キュー</span>
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
              <div><span>想定時間</span><strong>教材側で確認</strong></div>
              <div><span>テンプレート</span><strong>{template.name}</strong></div>
            </div>
            <div className="button-row answer-primary-actions">
              <button onClick={onSave}>一時保存</button>
              <button className="accent" onClick={onGoGrading}>採点へ進む</button>
              <button className="secondary" onClick={onNextProblem}>次の問題へ</button>
            </div>
          </section>

          <section className="focused-answer-area answer-input-main answer-only-main">
            <div className="focused-answer-title">
              <div>
                <p className="eyebrow">答案入力</p>
                <h2>{template.name}</h2>
              </div>
              <button className="secondary" onClick={onApplyRowPoints}>行別得点合計を反映</button>
            </div>
            <AnswerTable answer={answer} template={template} onChange={onChange} onApplyRowPoints={onApplyRowPoints} />
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
