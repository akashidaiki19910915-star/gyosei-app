import { useEffect, useMemo, useState } from 'react';
import type { AnswerState, GradingStatus, MaterialPdf, MaterialPdfMapping, PracticeSession, ProblemDefinition, ReviewRank, ReviewState } from '../types';
import { formatDateTime } from '../utils/dates';
import { rankFromSelfGrading, updateReviewStateFromGrading } from '../utils/reviewScheduler';
import { PdfViewerPanel } from './PdfViewerPanel';

type PdfViewMode = 'problem' | 'answer' | 'explanation';
type StudyPanelMode = 'study' | 'grading';

interface SelfGradeInput {
  status: GradingStatus;
  score: number;
  maxScore: number;
  memo: string;
}

interface Props {
  problem: ProblemDefinition;
  answer: AnswerState;
  pdfs: MaterialPdf[];
  mappings: MaterialPdfMapping[];
  activeSession: PracticeSession | null;
  reviewState?: ReviewState;
  panelMode?: StudyPanelMode;
  onStartCurrentSession: () => Promise<PracticeSession>;
  onSessionChange: (session: PracticeSession) => Promise<void>;
  onSelfGrade: (input: SelfGradeInput) => Promise<void>;
  onRequestGrading?: () => void;
}

function pageRange(mapping: MaterialPdfMapping, mode: PdfViewMode): { start: number; end: number } | null {
  if (mode === 'problem') return { start: mapping.problemPageStart, end: mapping.problemPageEnd };
  if (mode === 'answer' && mapping.answerPageStart) return { start: mapping.answerPageStart, end: mapping.answerPageEnd ?? mapping.answerPageStart };
  if (mode === 'explanation' && mapping.explanationPageStart) return { start: mapping.explanationPageStart, end: mapping.explanationPageEnd ?? mapping.explanationPageStart };
  return null;
}

function viewLabel(mode: PdfViewMode): string {
  if (mode === 'answer') return '解答';
  if (mode === 'explanation') return '解説';
  return '問題';
}

export function PdfStudyPanel({ problem, answer, pdfs, mappings, activeSession, reviewState, panelMode = 'study', onStartCurrentSession, onSessionChange, onSelfGrade, onRequestGrading }: Props) {
  const mapping = useMemo(() => mappings.find((item) => item.problemId === problem.id), [mappings, problem.id]);
  const pdf = useMemo(() => pdfs.find((item) => item.id === mapping?.materialPdfId), [pdfs, mapping?.materialPdfId]);
  const [viewMode, setViewMode] = useState<PdfViewMode>(panelMode === 'grading' ? 'answer' : 'problem');
  const [page, setPage] = useState(1);
  const [showPdf, setShowPdf] = useState(panelMode === 'grading');
  const [showSelfGrade, setShowSelfGrade] = useState(panelMode === 'grading');
  const [status, setStatus] = useState<GradingStatus>('正解');
  const [score, setScore] = useState(answer.score || 0);
  const [maxScore, setMaxScore] = useState(answer.maxScore || 20);
  const [memo, setMemo] = useState(answer.reviewMemo || '');

  const safeViewMode = panelMode === 'study' ? 'problem' : viewMode;
  const range = mapping ? pageRange(mapping, safeViewMode) : null;
  const proposedRank: ReviewRank = rankFromSelfGrading(status, score, maxScore);
  const proposedReview = updateReviewStateFromGrading({ problemId: problem.id, previous: reviewState, status, score, maxScore });

  useEffect(() => {
    const nextMode: PdfViewMode = panelMode === 'grading' && mapping?.answerPageStart ? 'answer' : 'problem';
    setViewMode(nextMode);
    setShowPdf(panelMode === 'grading');
    setShowSelfGrade(panelMode === 'grading');
    const nextRange = mapping ? pageRange(mapping, nextMode) : null;
    setPage(nextRange?.start ?? 1);
  }, [problem.id, mapping?.id, panelMode]);

  async function ensureSession(): Promise<PracticeSession> {
    if (activeSession && activeSession.problemId === problem.id && !activeSession.completed) return activeSession;
    return onStartCurrentSession();
  }

  async function openView(nextMode: PdfViewMode) {
    if (!mapping) return;
    if (panelMode === 'study' && nextMode !== 'problem') return;
    if (nextMode !== 'problem') {
      const session = await ensureSession();
      await onSessionChange({
        ...session,
        openedAnswer: session.openedAnswer || nextMode === 'answer',
        openedExplanation: session.openedExplanation || nextMode === 'explanation',
      });
    }
    const nextRange = pageRange(mapping, nextMode);
    if (!nextRange) return;
    setViewMode(nextMode);
    setPage(nextRange.start);
    setShowPdf(true);
  }

  async function answerNow() {
    const session = await ensureSession();
    const submittedAt = new Date().toISOString();
    await onSessionChange({ ...session, submittedAt, openedAnswer: false, answerSnapshot: answer });
    onRequestGrading?.();
  }

  async function submitSelfGrade() {
    await onSelfGrade({ status, score, maxScore, memo });
    setShowSelfGrade(false);
  }

  if (!mapping || !pdf) {
    return (
      <section className="panel pdf-study-panel mode-card">
        <h2>問題PDF</h2>
        <p className="empty">この問題IDにはPDFページ紐付けがありません。「教材を設定する」からPDFを開き、問題・解答・解説ページを登録してください。</p>
        <button onClick={() => { void onStartCurrentSession(); }}>PDFなしで演習セッション開始</button>
      </section>
    );
  }

  return (
    <section className={`pdf-study-shell pdf-${panelMode}`}>
      <div className="panel pdf-study-panel mode-card pdf-launch-panel">
        <div className="section-heading">
          <div>
            <p className="eyebrow">現在表示：{viewLabel(safeViewMode)}ページ</p>
            <h2>{panelMode === 'grading' ? '解答・解説PDF' : '問題PDF'}：{pdf.title}</h2>
            <p>{problem.sectionLabel} {problem.displayId} / {problem.topic} / 想定 {mapping.estimatedMinutes} / 難易度 {mapping.difficulty}</p>
          </div>
          <div className="button-row">
            {panelMode === 'study' && <button className="accent" onClick={() => { void openView('problem'); }}>問題PDFを大きく開く</button>}
            {panelMode === 'grading' && <button className={viewMode === 'problem' ? 'accent' : 'secondary'} onClick={() => { void openView('problem'); }}>問題</button>}
            {panelMode === 'grading' && <button className={viewMode === 'answer' ? 'accent' : 'secondary'} disabled={!mapping.answerPageStart} onClick={() => { void openView('answer'); }}>解答</button>}
            {panelMode === 'grading' && <button className={viewMode === 'explanation' ? 'accent' : 'secondary'} disabled={!mapping.explanationPageStart} onClick={() => { void openView('explanation'); }}>解説</button>}
            {showPdf && <button className="secondary" onClick={() => setShowPdf(false)}>PDFを閉じる</button>}
          </div>
        </div>
        {activeSession && activeSession.problemId === problem.id && !activeSession.completed && <p className="ok-text">演習中：{activeSession.selectedTimeMode} / 開始 {formatDateTime(activeSession.startedAt)}</p>}
        {panelMode === 'study' && (
          <div className="study-actions">
            <button className="resume-button" onClick={() => { void answerNow(); }}>回答する / 自己採点へ進む</button>
          </div>
        )}
      </div>

      {showPdf && range && (
        <PdfViewerPanel
          pdf={pdf}
          title={`${problem.displayId} ${problem.topic}`}
          label={viewLabel(safeViewMode)}
          page={page}
          pageStart={range.start}
          pageEnd={range.end}
          large
          onPageChange={setPage}
          onClose={() => setShowPdf(false)}
        />
      )}

      {panelMode === 'grading' && showSelfGrade && (
        <div className="panel self-grade-panel">
          <h3>自己採点</h3>
          <div className="form-grid four-columns">
            <label>結果
              <select value={status} onChange={(event) => setStatus(event.target.value as GradingStatus)}>
                <option value="正解">正解</option>
                <option value="部分正解">部分正解</option>
                <option value="不正解">不正解</option>
                <option value="迷いあり">正解だが迷いあり</option>
              </select>
            </label>
            <label>得点<input type="number" value={score} onChange={(event) => setScore(Number(event.target.value) || 0)} /></label>
            <label>満点<input type="number" value={maxScore} onChange={(event) => setMaxScore(Number(event.target.value) || 0)} /></label>
            <label>自動提案<input readOnly value={`${proposedRank} / 次回 ${proposedReview.nextReviewDate}`} /></label>
          </div>
          <label>自己採点メモ<textarea value={memo} onChange={(event) => setMemo(event.target.value)} placeholder="解答解説を見て、自分のミス原因・次回注意点を短く記録" /></label>
          <button className="accent" onClick={() => { void submitSelfGrade(); }}>自己採点を履歴へ保存</button>
        </div>
      )}
    </section>
  );
}
