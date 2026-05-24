import { useEffect, useMemo, useRef, useState } from 'react';
import type { AnswerState, GradingStatus, MaterialPdf, MaterialPdfMapping, PracticeSession, ProblemDefinition, ReviewRank, ReviewState } from '../types';
import { formatDateTime } from '../utils/dates';
import { renderPdfPageToCanvas } from '../utils/pdfRenderer';
import { rankFromSelfGrading, updateReviewStateFromGrading } from '../utils/reviewScheduler';

type PdfViewMode = 'problem' | 'answer' | 'explanation';

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
  onStartCurrentSession: () => Promise<PracticeSession>;
  onSessionChange: (session: PracticeSession) => Promise<void>;
  onSelfGrade: (input: SelfGradeInput) => Promise<void>;
}

function pageRange(mapping: MaterialPdfMapping, mode: PdfViewMode): { start: number; end: number } | null {
  if (mode === 'problem') return { start: mapping.problemPageStart, end: mapping.problemPageEnd };
  if (mode === 'answer' && mapping.answerPageStart) return { start: mapping.answerPageStart, end: mapping.answerPageEnd ?? mapping.answerPageStart };
  if (mode === 'explanation' && mapping.explanationPageStart) return { start: mapping.explanationPageStart, end: mapping.explanationPageEnd ?? mapping.explanationPageStart };
  return null;
}

function clamp(value: number, start: number, end: number): number {
  return Math.max(start, Math.min(end, value));
}

export function PdfStudyPanel({ problem, answer, pdfs, mappings, activeSession, reviewState, onStartCurrentSession, onSessionChange, onSelfGrade }: Props) {
  const mapping = useMemo(() => mappings.find((item) => item.problemId === problem.id), [mappings, problem.id]);
  const pdf = useMemo(() => pdfs.find((item) => item.id === mapping?.materialPdfId), [pdfs, mapping?.materialPdfId]);
  const [viewMode, setViewMode] = useState<PdfViewMode>('problem');
  const [page, setPage] = useState(1);
  const [scale, setScale] = useState(1.15);
  const [loading, setLoading] = useState(false);
  const [renderError, setRenderError] = useState('');
  const [showSelfGrade, setShowSelfGrade] = useState(false);
  const [status, setStatus] = useState<GradingStatus>('正解');
  const [score, setScore] = useState(answer.score || 0);
  const [maxScore, setMaxScore] = useState(answer.maxScore || 20);
  const [memo, setMemo] = useState('');
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  const range = mapping ? pageRange(mapping, viewMode) : null;
  const proposedRank: ReviewRank = rankFromSelfGrading(status, score, maxScore);
  const proposedReview = updateReviewStateFromGrading({ problemId: problem.id, previous: reviewState, status, score, maxScore });

  useEffect(() => {
    setViewMode('problem');
    setShowSelfGrade(false);
    const nextRange = mapping ? pageRange(mapping, 'problem') : null;
    setPage(nextRange?.start ?? 1);
  }, [problem.id, mapping?.id]);

  useEffect(() => {
    if (!pdf || !canvasRef.current || !range) return;
    let cancelled = false;
    setLoading(true);
    setRenderError('');
    renderPdfPageToCanvas(pdf.pdfBlob, clamp(page, range.start, range.end), scale, canvasRef.current)
      .catch(() => {
        if (!cancelled) setRenderError('PDFページの表示に失敗しました。ファイル破損、保護PDF、またはブラウザ非対応の可能性があります。');
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => { cancelled = true; };
  }, [pdf, page, scale, range?.start, range?.end]);

  async function ensureSession(): Promise<PracticeSession> {
    if (activeSession && activeSession.problemId === problem.id && !activeSession.completed) return activeSession;
    return onStartCurrentSession();
  }

  async function openView(nextMode: PdfViewMode) {
    if (!mapping) return;
    if (nextMode !== 'problem') {
      if (!window.confirm('解答・解説を開きます。自己採点時のみ開いてください。')) return;
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
  }

  async function answerNow() {
    const session = await ensureSession();
    const submittedAt = new Date().toISOString();
    await onSessionChange({ ...session, submittedAt, openedAnswer: true, answerSnapshot: answer });
    setShowSelfGrade(true);
    await openView('answer');
  }

  async function submitSelfGrade() {
    await onSelfGrade({ status, score, maxScore, memo });
    setShowSelfGrade(false);
  }

  if (!mapping || !pdf) {
    return (
      <section className="panel pdf-study-panel">
        <h2>問題PDF</h2>
        <p className="empty">この問題IDにはPDFページ紐付けがありません。PDF教材VaultにPDFを保存し、PDFページ紐付けから問題ページを登録してください。</p>
        <button onClick={() => { void onStartCurrentSession(); }}>PDFなしで演習セッション開始</button>
      </section>
    );
  }

  return (
    <section className="panel pdf-study-panel">
      <div className="section-heading">
        <div>
          <h2>問題PDF：{pdf.title}</h2>
          <p>{problem.sectionLabel} {problem.displayId} / {problem.topic} / 想定 {mapping.estimatedMinutes} / 難易度 {mapping.difficulty}</p>
        </div>
        <div className="button-row">
          <button className={viewMode === 'problem' ? 'accent' : 'secondary'} onClick={() => { void openView('problem'); }}>問題</button>
          <button className={viewMode === 'answer' ? 'accent' : 'secondary'} disabled={!mapping.answerPageStart} onClick={() => { void openView('answer'); }}>解答</button>
          <button className={viewMode === 'explanation' ? 'accent' : 'secondary'} disabled={!mapping.explanationPageStart} onClick={() => { void openView('explanation'); }}>解説</button>
        </div>
      </div>
      <div className="pdf-toolbar">
        <button onClick={() => range && setPage(clamp(page - 1, range.start, range.end))}>前ページ</button>
        <label>ページ
          <input type="number" min={range?.start ?? 1} max={range?.end ?? pdf.pageCount} value={page} onChange={(event) => range && setPage(clamp(Number(event.target.value) || range.start, range.start, range.end))} />
        </label>
        <span>{range ? `${range.start}〜${range.end}ページ範囲` : 'ページ範囲なし'}</span>
        <button onClick={() => range && setPage(clamp(page + 1, range.start, range.end))}>次ページ</button>
        <button onClick={() => setScale((current) => Math.max(0.7, current - 0.15))}>縮小</button>
        <button onClick={() => setScale((current) => Math.min(2.2, current + 0.15))}>拡大</button>
      </div>
      {activeSession && activeSession.problemId === problem.id && !activeSession.completed && <p className="ok-text">演習中：{activeSession.selectedTimeMode} / 開始 {formatDateTime(activeSession.startedAt)}</p>}
      {renderError && <p className="warning-text">{renderError}</p>}
      <div className="pdf-canvas-wrap">
        {loading && <p className="empty">PDFを表示中...</p>}
        <canvas ref={canvasRef} className="pdf-canvas" />
      </div>
      <div className="study-actions">
        <button className="resume-button" onClick={() => { void answerNow(); }}>回答する / 自己採点へ進む</button>
      </div>
      {showSelfGrade && (
        <div className="self-grade-panel">
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
