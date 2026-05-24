import { useEffect, useMemo, useState } from 'react';
import { problemCatalog } from '../data/problemCatalog';
import type { MaterialPdf, MaterialPdfMapping, PdfProblemExtractionCandidate, TemplateId } from '../types';
import { nowIso } from '../utils/dates';
import { extractProblemCandidatesFromPdf } from '../utils/pdfAutoLink';
import { formatFileSize, getPdfPageCount } from '../utils/pdfRenderer';
import { PdfViewerPanel } from './PdfViewerPanel';

interface Props {
  pdfs: MaterialPdf[];
  mappings: MaterialPdfMapping[];
  selectedProblemId: string;
  onSavePdf: (pdf: MaterialPdf) => Promise<void>;
  onDeletePdf: (id: string) => Promise<void>;
  onSaveMapping: (mapping: MaterialPdfMapping) => Promise<void>;
  onDeleteMapping: (id: string) => Promise<void>;
  onOpenProblem: (problemId: string) => void | Promise<void>;
  onMessage: (message: string) => void;
}

function emptyCandidate(pdfId: string): PdfProblemExtractionCandidate {
  const problem = problemCatalog[0];
  return {
    id: crypto.randomUUID(),
    materialPdfId: pdfId,
    suggestedProblemId: problem.id,
    displayId: problem.displayId,
    subject: problem.subject,
    sectionId: problem.sectionId,
    sectionLabel: problem.sectionLabel,
    topic: problem.topic,
    templateId: problem.defaultTemplateId,
    problemKind: problem.defaultTemplateId === 'journal' ? 'journal' : 'calculationTable',
    problemText: '',
    answerText: '',
    explanationText: '',
    sourcePages: [],
    answerPages: [],
    explanationPages: [],
    confidence: 0,
    status: 'needsReview',
    reason: '手動で問題IDと抽出テキストを確認してください。',
  };
}

function pagesToRange(pages: number[], fallback = 1): { start: number; end: number } {
  if (pages.length === 0) return { start: fallback, end: fallback };
  const sorted = [...pages].sort((a, b) => a - b);
  return { start: sorted[0], end: sorted[sorted.length - 1] };
}

function candidateToMapping(candidate: PdfProblemExtractionCandidate): MaterialPdfMapping {
  const now = nowIso();
  const problem = problemCatalog.find((item) => item.id === candidate.suggestedProblemId) ?? problemCatalog[0];
  const problemRange = pagesToRange(candidate.sourcePages);
  const answerRange = pagesToRange(candidate.answerPages, problemRange.end);
  const explanationRange = pagesToRange(candidate.explanationPages, answerRange.end);
  return {
    id: `${candidate.materialPdfId}-${problem.id}`,
    qualificationId: 'nissho_boki2',
    problemId: problem.id,
    displayId: problem.displayId,
    materialPdfId: candidate.materialPdfId,
    subjectId: problem.subject,
    topicId: problem.displayId,
    problemKind: candidate.problemKind,
    problemText: candidate.problemText,
    answerText: candidate.answerText,
    explanationText: candidate.explanationText,
    sourcePages: candidate.sourcePages,
    answerPages: candidate.answerPages,
    explanationPages: candidate.explanationPages,
    extractionConfidence: candidate.confidence,
    extractionStatus: candidate.status,
    extractedAt: now,
    problemPageStart: problemRange.start,
    problemPageEnd: problemRange.end,
    answerPageStart: candidate.answerPages.length ? answerRange.start : undefined,
    answerPageEnd: candidate.answerPages.length ? answerRange.end : undefined,
    explanationPageStart: candidate.explanationPages.length ? explanationRange.start : undefined,
    explanationPageEnd: candidate.explanationPages.length ? explanationRange.end : undefined,
    estimatedMinutes: '10分',
    difficulty: candidate.confidence >= 75 ? '標準' : '難',
    memo: candidate.reason,
    problemBlocks: [
      {
        id: `${problem.id}-text-main`,
        title: '設問1',
        description: problem.topic,
        questionText: candidate.problemText,
        sourcePages: candidate.sourcePages,
        problemPageStart: problemRange.start,
        problemPageEnd: problemRange.end,
        templateId: candidate.templateId,
        estimatedMinutes: 10,
        difficulty: '標準',
        memo: '',
      },
    ],
    createdAt: now,
    updatedAt: now,
  };
}

function statusLabel(status: PdfProblemExtractionCandidate['status']): string {
  if (status === 'autoLinked') return '自動紐づけ済み';
  if (status === 'needsReview') return '確認が必要';
  return '未分類';
}

function statusClass(status: PdfProblemExtractionCandidate['status']): string {
  if (status === 'autoLinked') return 'good';
  if (status === 'needsReview') return 'warn';
  return 'danger';
}

export function MaterialSetupWorkspace({ pdfs, mappings, selectedProblemId, onSavePdf, onDeletePdf, onSaveMapping, onDeleteMapping, onOpenProblem, onMessage }: Props) {
  const [title, setTitle] = useState('');
  const [sourceName, setSourceName] = useState('CPA問題集');
  const [bookType, setBookType] = useState<MaterialPdf['bookType']>('商業簿記');
  const [selectedPdfId, setSelectedPdfId] = useState(pdfs[0]?.id ?? '');
  const [candidates, setCandidates] = useState<PdfProblemExtractionCandidate[]>([]);
  const [selectedCandidateId, setSelectedCandidateId] = useState('');
  const [extracting, setExtracting] = useState(false);
  const [progress, setProgress] = useState('');
  const [previewPage, setPreviewPage] = useState(1);

  const selectedPdf = useMemo(() => pdfs.find((pdf) => pdf.id === selectedPdfId) ?? pdfs[0], [pdfs, selectedPdfId]);
  const selectedCandidate = candidates.find((candidate) => candidate.id === selectedCandidateId) ?? candidates[0];
  const counts = useMemo(() => ({
    autoLinked: candidates.filter((candidate) => candidate.status === 'autoLinked').length,
    needsReview: candidates.filter((candidate) => candidate.status === 'needsReview').length,
    unmatched: candidates.filter((candidate) => candidate.status === 'unmatched').length,
  }), [candidates]);
  const mappedProblemIds = useMemo(() => new Set(mappings.map((mapping) => mapping.problemId)), [mappings]);
  const unmappedCount = problemCatalog.filter((problem) => !mappedProblemIds.has(problem.id)).length;

  useEffect(() => {
    if (!selectedPdfId && pdfs[0]) setSelectedPdfId(pdfs[0].id);
  }, [pdfs, selectedPdfId]);

  useEffect(() => {
    const mapping = mappings.find((item) => item.problemId === selectedProblemId);
    if (mapping?.sourcePages?.[0]) setPreviewPage(mapping.sourcePages[0]);
  }, [mappings, selectedProblemId]);

  async function importPdf(file: File | undefined) {
    if (!file) return;
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      onMessage('PDFファイルのみ取り込めます');
      return;
    }
    try {
      const now = nowIso();
      const pageCount = await getPdfPageCount(file);
      const pdf: MaterialPdf = {
        id: crypto.randomUUID(),
        title: title.trim() || file.name.replace(/\.pdf$/i, ''),
        sourceName: sourceName.trim() || '教材PDF',
        examType: '日商簿記2級',
        bookType,
        fileName: file.name,
        mimeType: file.type || 'application/pdf',
        size: file.size,
        pageCount,
        pdfBlob: file,
        createdAt: now,
        updatedAt: now,
      };
      await onSavePdf(pdf);
      setSelectedPdfId(pdf.id);
      setTitle('');
      setCandidates([]);
      setSelectedCandidateId('');
      onMessage('PDFを端末内IndexedDBへ保存しました。次に「PDFから問題候補を抽出」を押してください。');
    } catch {
      onMessage('PDFの読み込みに失敗しました。破損ファイル、保護PDF、ブラウザ非対応の可能性があります。');
    }
  }

  async function extractCandidates() {
    if (!selectedPdf) {
      onMessage('先にPDFを選択してください');
      return;
    }
    setExtracting(true);
    setProgress('抽出準備中');
    try {
      const rows = await extractProblemCandidatesFromPdf(selectedPdf, (page, pageCount) => setProgress(`PDFテキスト抽出中 ${page} / ${pageCount}ページ`));
      setCandidates(rows);
      setSelectedCandidateId(rows[0]?.id ?? '');
      setProgress('');
      onMessage(`問題候補を抽出しました。自動紐づけ済み：${rows.filter((row) => row.status === 'autoLinked').length}件、確認が必要：${rows.filter((row) => row.status === 'needsReview').length}件、未分類：${rows.filter((row) => row.status === 'unmatched').length}件`);
    } catch {
      onMessage('PDFテキスト抽出に失敗しました。テキストPDFでない場合、OCRは今回実装していないため抽出できません。');
    } finally {
      setExtracting(false);
      setProgress('');
    }
  }

  function updateCandidate(id: string, patch: Partial<PdfProblemExtractionCandidate>) {
    setCandidates((current) => current.map((candidate) => {
      if (candidate.id !== id) return candidate;
      const nextProblemId = patch.suggestedProblemId ?? candidate.suggestedProblemId;
      const problem = problemCatalog.find((item) => item.id === nextProblemId) ?? problemCatalog[0];
      return {
        ...candidate,
        ...patch,
        suggestedProblemId: nextProblemId,
        displayId: problem.displayId,
        subject: problem.subject,
        sectionId: problem.sectionId,
        sectionLabel: problem.sectionLabel,
        topic: problem.topic,
        templateId: (patch.templateId ?? problem.defaultTemplateId) as TemplateId,
        problemKind: problem.defaultTemplateId === 'journal' ? 'journal' : 'calculationTable',
      };
    }));
  }

  async function saveCandidate(candidate: PdfProblemExtractionCandidate) {
    await onSaveMapping(candidateToMapping({ ...candidate, status: candidate.confidence >= 75 ? 'autoLinked' : candidate.status }));
    onMessage(`${candidate.displayId} の抽出テキスト紐づけを保存しました`);
  }

  async function saveAutoLinkedCandidates() {
    const targets = candidates.filter((candidate) => candidate.status === 'autoLinked');
    if (targets.length === 0) {
      onMessage('自動保存できる高信頼度候補がありません。確認が必要な候補を手動保存してください。');
      return;
    }
    for (const candidate of targets) await onSaveMapping(candidateToMapping(candidate));
    onMessage(`自動紐づけ済み ${targets.length}件を保存しました。確認が必要な候補だけ手動修正してください。`);
  }

  function addManualCandidate() {
    if (!selectedPdf) return;
    const next = emptyCandidate(selectedPdf.id);
    setCandidates((current) => [next, ...current]);
    setSelectedCandidateId(next.id);
  }

  return (
    <section className="material-text-setup">
      <section className="panel material-setup-main">
        <div className="section-heading">
          <div>
            <p className="eyebrow">教材を設定する</p>
            <h2>PDF取込・テキスト抽出・問題ID自動紐づけ</h2>
            <p>通常演習画面へ出す問題文は、PDF画像ではなく端末内で抽出したテキストを使います。PDF本体・抽出テキストはGitHubへ保存しません。</p>
          </div>
          <button className="secondary" onClick={() => { void onOpenProblem(selectedProblemId); }}>今すぐ解くへ戻る</button>
        </div>

        <div className="setup-step-grid">
          <section className="setup-step-card">
            <h3>1. PDFを選択する</h3>
            <div className="form-grid three-columns">
              <label>表示名<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="例：商業簿記 試験対策編" /></label>
              <label>教材区分<select value={bookType} onChange={(event) => setBookType(event.target.value as MaterialPdf['bookType'])}><option value="商業簿記">商業簿記</option><option value="工業簿記">工業簿記</option><option value="その他">その他</option></select></label>
              <label>出所メモ<input value={sourceName} onChange={(event) => setSourceName(event.target.value)} placeholder="例：正規取得PDF" /></label>
            </div>
            <label className="import-label pdf-import-label">PDFを選択<input type="file" accept="application/pdf" onChange={(event) => { void importPdf(event.target.files?.[0]); }} /></label>
            <label>対象PDF<select value={selectedPdfId} onChange={(event) => { setSelectedPdfId(event.target.value); setCandidates([]); setSelectedCandidateId(''); }}><option value="">PDFを選択</option>{pdfs.map((pdf) => <option key={pdf.id} value={pdf.id}>{pdf.title}（{pdf.pageCount}P / {formatFileSize(pdf.size)}）</option>)}</select></label>
          </section>

          <section className="setup-step-card">
            <h3>2. PDFから問題候補を一括抽出する</h3>
            <p>「第1問対策」「1-1」「解答・解説」などの見出し・問題ID・論点語から候補を作ります。信頼度が低いものだけ手動確認します。</p>
            <div className="button-row">
              <button className="accent" disabled={!selectedPdf || extracting} onClick={() => { void extractCandidates(); }}>{extracting ? '抽出中...' : 'PDFから問題候補を抽出'}</button>
              <button className="secondary" disabled={!selectedPdf} onClick={addManualCandidate}>手動候補を追加</button>
            </div>
            {progress && <p className="notice-small">{progress}</p>}
          </section>

          <section className="setup-step-card extraction-summary-card">
            <h3>3. 自動紐づけ結果</h3>
            <div className="progress-summary-grid extraction-counts">
              <div><strong>{counts.autoLinked}</strong><span>自動紐づけ済み</span></div>
              <div><strong>{counts.needsReview}</strong><span>確認が必要</span></div>
              <div><strong>{counts.unmatched}</strong><span>未分類</span></div>
              <div><strong>{unmappedCount}</strong><span>未設定問題ID</span></div>
            </div>
            <button className="accent" disabled={counts.autoLinked === 0} onClick={() => { void saveAutoLinkedCandidates(); }}>高信頼度候補を一括保存</button>
          </section>
        </div>
      </section>

      {candidates.length > 0 && <section className="panel candidate-review-panel">
        <div className="section-heading">
          <div><p className="eyebrow">候補確認</p><h2>確認が必要な問題だけ手動修正する</h2></div>
          {selectedCandidate && <button className="accent" onClick={() => { void saveCandidate(selectedCandidate); }}>この候補を保存</button>}
        </div>
        <div className="candidate-review-layout">
          <div className="candidate-list">
            {candidates.map((candidate) => <button key={candidate.id} className={`candidate-list-item ${selectedCandidate?.id === candidate.id ? 'active' : ''}`} onClick={() => { setSelectedCandidateId(candidate.id); setPreviewPage(candidate.sourcePages[0] ?? 1); }}><strong>{candidate.displayId}</strong><span className={`status-pill ${statusClass(candidate.status)}`}>{statusLabel(candidate.status)}</span><span>信頼度 {candidate.confidence}%</span><small>{candidate.topic}</small></button>)}
          </div>

          {selectedCandidate && <div className="candidate-editor">
            <div className="form-grid three-columns">
              <label>問題ID<select value={selectedCandidate.suggestedProblemId} onChange={(event) => updateCandidate(selectedCandidate.id, { suggestedProblemId: event.target.value })}>{problemCatalog.map((problem) => <option key={problem.id} value={problem.id}>{problem.displayId} {problem.topic}</option>)}</select></label>
              <label>状態<select value={selectedCandidate.status} onChange={(event) => updateCandidate(selectedCandidate.id, { status: event.target.value as PdfProblemExtractionCandidate['status'] })}><option value="autoLinked">自動紐づけ済み</option><option value="needsReview">確認が必要</option><option value="unmatched">未分類</option></select></label>
              <label>信頼度<input type="number" min="0" max="100" value={selectedCandidate.confidence} onChange={(event) => updateCandidate(selectedCandidate.id, { confidence: Number(event.target.value) || 0 })} /></label>
            </div>
            <div className="form-grid three-columns">
              <label>問題ページ<input value={selectedCandidate.sourcePages.join(',')} onChange={(event) => updateCandidate(selectedCandidate.id, { sourcePages: event.target.value.split(',').map((item) => Number(item.trim())).filter(Boolean) })} /></label>
              <label>解答ページ<input value={selectedCandidate.answerPages.join(',')} onChange={(event) => updateCandidate(selectedCandidate.id, { answerPages: event.target.value.split(',').map((item) => Number(item.trim())).filter(Boolean) })} /></label>
              <label>解説ページ<input value={selectedCandidate.explanationPages.join(',')} onChange={(event) => updateCandidate(selectedCandidate.id, { explanationPages: event.target.value.split(',').map((item) => Number(item.trim())).filter(Boolean) })} /></label>
            </div>
            <label>問題文テキスト<textarea className="large-textarea" value={selectedCandidate.problemText} onChange={(event) => updateCandidate(selectedCandidate.id, { problemText: event.target.value })} /></label>
            <label>解答テキスト<textarea className="large-textarea" value={selectedCandidate.answerText} onChange={(event) => updateCandidate(selectedCandidate.id, { answerText: event.target.value })} /></label>
            <label>解説テキスト<textarea className="large-textarea" value={selectedCandidate.explanationText} onChange={(event) => updateCandidate(selectedCandidate.id, { explanationText: event.target.value })} /></label>
          </div>}
        </div>
      </section>}

      <section className="material-setup-grid compact-material-grid">
        <div className="setup-preview-panel"><PdfViewerPanel pdf={selectedPdf} title={selectedPdf ? selectedPdf.title : 'PDF未選択'} label="原本確認" page={previewPage} pageStart={1} pageEnd={selectedPdf?.pageCount ?? 1} onPageChange={setPreviewPage} large /></div>
        <details className="panel mapping-list-panel" open><summary>保存済み設定</summary><div className="panel-body"><div className="table-wrap short-wrap"><table className="compact-table"><thead><tr><th>問題ID</th><th>論点</th><th>PDF</th><th>抽出</th><th>信頼度</th><th>操作</th></tr></thead><tbody>{mappings.map((mapping) => { const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0]; const pdf = pdfs.find((item) => item.id === mapping.materialPdfId); return <tr key={mapping.id}><td>{problem.displayId}</td><td>{problem.topic}</td><td>{pdf?.title ?? 'PDF未登録'}</td><td>{mapping.problemText ? 'テキストあり' : '未抽出'}</td><td>{mapping.extractionConfidence ?? '-'}</td><td><button onClick={() => { void onOpenProblem(mapping.problemId); }}>開く</button> <button className="danger" onClick={() => { void onDeleteMapping(mapping.id); }}>削除</button></td></tr>; })}{mappings.length === 0 && <tr><td colSpan={6}>まだ保存済み設定がありません。</td></tr>}</tbody></table></div></div></details>
        <details className="panel mapping-list-panel"><summary>PDF教材データの削除</summary><div className="panel-body"><div className="table-wrap short-wrap"><table className="compact-table"><thead><tr><th>教材名</th><th>ファイル</th><th>ページ数</th><th>サイズ</th><th>操作</th></tr></thead><tbody>{pdfs.map((pdf) => <tr key={pdf.id}><td>{pdf.title}</td><td>{pdf.fileName}</td><td>{pdf.pageCount}</td><td>{formatFileSize(pdf.size)}</td><td><button className="danger" onClick={() => { if (window.confirm('このPDF教材データと問題IDへの紐付けを削除します。答案履歴・白紙再現カードは削除されません。本当に削除しますか？')) void onDeletePdf(pdf.id); }}>削除</button></td></tr>)}{pdfs.length === 0 && <tr><td colSpan={5}>PDF教材は未登録です。</td></tr>}</tbody></table></div></div></details>
      </section>
    </section>
  );
}
