import { useEffect, useMemo, useState } from 'react';
import { problemCatalog } from '../data/problemCatalog';
import { templateCatalog } from '../data/templateCatalog';
import type { BlockDifficulty, MaterialPdf, MaterialPdfMapping, ProblemBlock, TemplateId, TimeMode } from '../types';
import { nowIso } from '../utils/dates';
import { timeModes } from '../utils/disposableStudy';
import { suggestPdfLinkCandidates, type PdfAutoLinkCandidate } from '../utils/pdfAutoLink';
import { formatFileSize, getPdfPageCount } from '../utils/pdfRenderer';
import { PdfQuestionPreview } from './PdfQuestionPreview';
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

type RangeKind = 'problem' | 'answer' | 'explanation';

function blankBlock(index: number, templateId: TemplateId = 'journal'): ProblemBlock {
  const start = Math.min(90, index * 30);
  return {
    id: crypto.randomUUID(),
    title: `設問${index + 1}`,
    description: '',
    problemPageStart: 1,
    problemPageEnd: 1,
    cropTopPercent: start,
    cropBottomPercent: Math.min(100, start + 30),
    estimatedMinutes: 3,
    difficulty: '標準',
    templateId,
    memo: '',
  };
}

function blankMapping(problemId: string, pdfId = ''): MaterialPdfMapping {
  const problem = problemCatalog.find((item) => item.id === problemId) ?? problemCatalog[0];
  const now = nowIso();
  return {
    id: crypto.randomUUID(),
    problemId: problem.id,
    displayId: problem.displayId,
    materialPdfId: pdfId,
    problemPageStart: 1,
    problemPageEnd: 1,
    answerPageStart: undefined,
    answerPageEnd: undefined,
    explanationPageStart: undefined,
    explanationPageEnd: undefined,
    estimatedMinutes: '10分',
    difficulty: '標準',
    memo: '',
    problemBlocks: [blankBlock(0, problem.defaultTemplateId)],
    createdAt: now,
    updatedAt: now,
  };
}

function numberOrUndefined(value: string): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function percentOrDefault(value: string, fallback: number): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.min(100, parsed));
}

function questionTitle(index: number, title?: string): string {
  const trimmed = title?.trim();
  if (!trimmed || trimmed.startsWith('問題文')) return `設問${index + 1}`;
  return trimmed;
}

function ensureBlocks(mapping: MaterialPdfMapping): ProblemBlock[] {
  if (mapping.problemBlocks && mapping.problemBlocks.length > 0) {
    return mapping.problemBlocks.map((block, index) => ({
      ...block,
      title: questionTitle(index, block.title),
      problemPageStart: block.problemPageStart ?? mapping.problemPageStart,
      problemPageEnd: block.problemPageEnd ?? mapping.problemPageEnd,
      cropTopPercent: block.cropTopPercent ?? 0,
      cropBottomPercent: block.cropBottomPercent ?? 100,
    }));
  }
  const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0];
  return [blankBlock(0, problem.defaultTemplateId)];
}

function blockDifficultyLabel(value?: BlockDifficulty): string {
  return value ?? '標準';
}

export function MaterialSetupWorkspace({ pdfs, mappings, selectedProblemId, onSavePdf, onDeletePdf, onSaveMapping, onDeleteMapping, onOpenProblem, onMessage }: Props) {
  const [title, setTitle] = useState('');
  const [sourceName, setSourceName] = useState('CPA問題集');
  const [bookType, setBookType] = useState<MaterialPdf['bookType']>('商業簿記');
  const [selectedPdfId, setSelectedPdfId] = useState(pdfs[0]?.id ?? '');
  const [draft, setDraft] = useState<MaterialPdfMapping>(() => blankMapping(selectedProblemId, pdfs[0]?.id ?? ''));
  const [previewKind, setPreviewKind] = useState<RangeKind>('problem');
  const [page, setPage] = useState(1);
  const [candidates, setCandidates] = useState<PdfAutoLinkCandidate[]>([]);
  const [scanning, setScanning] = useState(false);

  const selectedPdf = useMemo(() => pdfs.find((pdf) => pdf.id === selectedPdfId) ?? pdfs[0], [pdfs, selectedPdfId]);
  const mappedProblemIds = useMemo(() => new Set(mappings.map((mapping) => mapping.problemId)), [mappings]);
  const unmapped = problemCatalog.filter((problem) => !mappedProblemIds.has(problem.id));
  const mapped = problemCatalog.filter((problem) => mappedProblemIds.has(problem.id));
  const blocks = ensureBlocks(draft);

  useEffect(() => {
    if (!selectedPdfId && pdfs[0]) setSelectedPdfId(pdfs[0].id);
  }, [pdfs, selectedPdfId]);

  useEffect(() => {
    const existing = mappings.find((mapping) => mapping.problemId === selectedProblemId) ?? blankMapping(selectedProblemId, selectedPdf?.id ?? '');
    const next = { ...existing, materialPdfId: existing.materialPdfId || selectedPdf?.id || '', problemBlocks: ensureBlocks(existing) };
    setDraft(next);
    setSelectedPdfId(next.materialPdfId || selectedPdf?.id || '');
  }, [selectedProblemId, mappings, selectedPdf?.id]);

  const activeRange = (() => {
    if (previewKind === 'answer' && draft.answerPageStart) return { start: draft.answerPageStart, end: draft.answerPageEnd ?? draft.answerPageStart };
    if (previewKind === 'explanation' && draft.explanationPageStart) return { start: draft.explanationPageStart, end: draft.explanationPageEnd ?? draft.explanationPageStart };
    return { start: draft.problemPageStart, end: draft.problemPageEnd };
  })();

  useEffect(() => {
    setPage(activeRange.start || 1);
  }, [previewKind, draft.problemPageStart, draft.answerPageStart, draft.explanationPageStart]);

  function update(patch: Partial<MaterialPdfMapping>) {
    const nextProblemId = patch.problemId ?? draft.problemId;
    const problem = problemCatalog.find((item) => item.id === nextProblemId) ?? problemCatalog[0];
    setDraft((current) => ({ ...current, ...patch, problemId: nextProblemId, displayId: problem.displayId, problemBlocks: patch.problemBlocks ?? ensureBlocks(current), updatedAt: nowIso() }));
  }

  function updateBlock(index: number, patch: Partial<ProblemBlock>) {
    update({ problemBlocks: blocks.map((block, blockIndex) => blockIndex === index ? { ...block, ...patch } : block) });
  }

  function addBlock() {
    update({ problemBlocks: [...blocks, blankBlock(blocks.length, blocks[blocks.length - 1]?.templateId ?? 'journal')] });
  }

  function deleteBlock(index: number) {
    const next = blocks.filter((_, blockIndex) => blockIndex !== index);
    update({ problemBlocks: next.length > 0 ? next : [blankBlock(0)] });
  }

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
      setDraft((current) => ({ ...current, materialPdfId: pdf.id }));
      setTitle('');
      setCandidates([]);
      onMessage('PDFを端末内IndexedDBへ保存しました。続けて問題IDと設問ごとの表示範囲を登録してください。');
    } catch {
      onMessage('PDFの読み込みに失敗しました。破損ファイル、保護PDF、ブラウザ非対応の可能性があります。');
    }
  }

  async function scanCandidates() {
    if (!selectedPdf) return;
    setScanning(true);
    try {
      const rows = await suggestPdfLinkCandidates(selectedPdf);
      setCandidates(rows);
      onMessage(rows.length ? `候補ページを${rows.length}件検出しました。必ずPDFを確認してから保存してください。` : '候補ページは検出できませんでした。PDFを見ながら手動で指定してください。');
    } finally {
      setScanning(false);
    }
  }

  async function saveMapping() {
    if (!draft.materialPdfId) {
      onMessage('先に対象PDFを選択してください');
      return;
    }
    await onSaveMapping({ ...draft, problemBlocks: blocks, updatedAt: nowIso() });
  }

  function jumpToRange(kind: RangeKind) {
    setPreviewKind(kind);
    if (kind === 'answer' && draft.answerPageStart) setPage(draft.answerPageStart);
    else if (kind === 'explanation' && draft.explanationPageStart) setPage(draft.explanationPageStart);
    else setPage(draft.problemPageStart);
  }

  function applyCandidate(pageNumber: number, target: RangeKind) {
    if (target === 'problem') update({ problemPageStart: pageNumber, problemPageEnd: Math.max(pageNumber, draft.problemPageEnd) });
    if (target === 'answer') update({ answerPageStart: pageNumber, answerPageEnd: pageNumber });
    if (target === 'explanation') update({ explanationPageStart: pageNumber, explanationPageEnd: pageNumber });
    setPage(pageNumber);
  }

  return (
    <section className="material-setup-grid wizard-setup-grid">
      <div className="panel setup-control-panel mode-card wizard-control-panel">
        <p className="eyebrow">教材を設定する</p>
        <h2>PDF教材を選ぶ → 問題IDを選ぶ → 設問ごとの表示範囲を登録</h2>
        <p className="notice-small">PDF本文をアプリに保存するのではなく、購入済みPDFの該当ページ範囲だけを端末内で参照します。</p>
        <p className="notice-small">PDFの作りによっては自動候補が出ない場合があります。その場合はPDFを見ながら手動でページ範囲を指定してください。</p>

        <details open className="wizard-step-card"><summary>1. PDF教材を選ぶ</summary>
          <div className="form-grid three-columns">
            <label>表示名<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="例：商業簿記 試験対策編" /></label>
            <label>教材区分<select value={bookType} onChange={(event) => setBookType(event.target.value as MaterialPdf['bookType'])}><option value="商業簿記">商業簿記</option><option value="工業簿記">工業簿記</option><option value="その他">その他</option></select></label>
            <label>出所メモ<input value={sourceName} onChange={(event) => setSourceName(event.target.value)} placeholder="例：正規取得PDF" /></label>
          </div>
          <label className="import-label pdf-import-label">PDFを選択してプレビュー<input type="file" accept="application/pdf" onChange={(event) => { void importPdf(event.target.files?.[0]); }} /></label>
        </details>

        <details open className="wizard-step-card"><summary>2. 問題IDとページを選ぶ</summary>
          <div className="form-grid two-columns"><label>対象PDF<select value={draft.materialPdfId || selectedPdfId} onChange={(event) => { setSelectedPdfId(event.target.value); update({ materialPdfId: event.target.value }); }}><option value="">PDFを選択</option>{pdfs.map((pdf) => <option key={pdf.id} value={pdf.id}>{pdf.title}（{pdf.pageCount}P）</option>)}</select></label><label>問題ID<select value={draft.problemId} onChange={(event) => { const existing = mappings.find((mapping) => mapping.problemId === event.target.value); setDraft(existing ? { ...existing, problemBlocks: ensureBlocks(existing) } : blankMapping(event.target.value, draft.materialPdfId || selectedPdfId)); }}>
            {problemCatalog.map((problem) => <option key={problem.id} value={problem.id}>{problem.sectionLabel} {problem.displayId} {problem.topic}</option>)}</select></label></div>
          <div className="form-grid six-columns"><label>問題開始P<input type="number" min="1" value={draft.problemPageStart} onChange={(event) => update({ problemPageStart: Number(event.target.value) || 1 })} /></label><label>問題終了P<input type="number" min="1" value={draft.problemPageEnd} onChange={(event) => update({ problemPageEnd: Number(event.target.value) || draft.problemPageStart })} /></label><label>解答開始P<input type="number" min="1" value={draft.answerPageStart ?? ''} onChange={(event) => update({ answerPageStart: numberOrUndefined(event.target.value) })} /></label><label>解答終了P<input type="number" min="1" value={draft.answerPageEnd ?? ''} onChange={(event) => update({ answerPageEnd: numberOrUndefined(event.target.value) })} /></label><label>解説開始P<input type="number" min="1" value={draft.explanationPageStart ?? ''} onChange={(event) => update({ explanationPageStart: numberOrUndefined(event.target.value) })} /></label><label>解説終了P<input type="number" min="1" value={draft.explanationPageEnd ?? ''} onChange={(event) => update({ explanationPageEnd: numberOrUndefined(event.target.value) })} /></label></div>
          <div className="form-grid three-columns"><label>想定時間<select value={draft.estimatedMinutes} onChange={(event) => update({ estimatedMinutes: event.target.value as TimeMode })}>{timeModes.map((mode) => <option key={mode} value={mode}>{mode}</option>)}</select></label><label>難易度<select value={draft.difficulty} onChange={(event) => update({ difficulty: event.target.value as MaterialPdfMapping['difficulty'] })}><option value="易">易</option><option value="標準">標準</option><option value="難">難</option></select></label><label>メモ<input value={draft.memo} onChange={(event) => update({ memo: event.target.value })} placeholder="教材本文は入れない" /></label></div>
          <div className="button-row"><button onClick={() => { void scanCandidates(); }} disabled={!selectedPdf || scanning}>{scanning ? '候補抽出中...' : '候補ページを提案'}</button><button className={previewKind === 'problem' ? 'accent' : 'secondary'} onClick={() => jumpToRange('problem')}>問題</button><button className={previewKind === 'answer' ? 'accent' : 'secondary'} disabled={!draft.answerPageStart} onClick={() => jumpToRange('answer')}>解答</button><button className={previewKind === 'explanation' ? 'accent' : 'secondary'} disabled={!draft.explanationPageStart} onClick={() => jumpToRange('explanation')}>解説</button></div>
          {candidates.length > 0 && <div className="auto-link-candidates"><h3>候補ページ</h3>{candidates.map((candidate) => <div key={candidate.page} className="candidate-row"><strong>P{candidate.page}</strong><span>{candidate.labels.join(' / ')}</span><small>{candidate.preview}</small><button onClick={() => setPage(candidate.page)}>このページへ移動</button><button onClick={() => applyCandidate(candidate.page, 'problem')}>問題開始</button><button onClick={() => applyCandidate(candidate.page, 'answer')}>解答開始</button><button onClick={() => applyCandidate(candidate.page, 'explanation')}>解説開始</button></div>)}</div>}
        </details>

        <details open className="wizard-step-card"><summary>3. 設問ごとのPDF表示範囲と答案欄を作る</summary>
          <div className="button-row"><button className="accent" onClick={addBlock}>設問を追加</button><span>現在 {blocks.length} 設問</span></div>
          <div className="block-editor-list">
            {blocks.map((block, index) => <div key={block.id} className="block-editor-card"><div className="section-heading"><div><p className="eyebrow">設問{index + 1}</p><h3>{questionTitle(index, block.title)}</h3></div><button className="danger" onClick={() => deleteBlock(index)}>削除</button></div><div className="form-grid three-columns"><label>設問名<input value={questionTitle(index, block.title)} onChange={(event) => updateBlock(index, { title: event.target.value })} /></label><label>答案テンプレート<select value={block.templateId} onChange={(event) => updateBlock(index, { templateId: event.target.value as TemplateId })}>{templateCatalog.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}</select></label><label>想定分<input type="number" value={block.estimatedMinutes ?? ''} onChange={(event) => updateBlock(index, { estimatedMinutes: numberOrUndefined(event.target.value) })} /></label></div><div className="form-grid six-columns"><label>問題P<input type="number" value={block.problemPageStart ?? draft.problemPageStart} onChange={(event) => updateBlock(index, { problemPageStart: numberOrUndefined(event.target.value) })} /></label><label>問題終了P<input type="number" value={block.problemPageEnd ?? block.problemPageStart ?? draft.problemPageEnd} onChange={(event) => updateBlock(index, { problemPageEnd: numberOrUndefined(event.target.value) })} /></label><label>上から％<input type="number" min="0" max="100" value={block.cropTopPercent ?? 0} onChange={(event) => updateBlock(index, { cropTopPercent: percentOrDefault(event.target.value, 0) })} /></label><label>下まで％<input type="number" min="0" max="100" value={block.cropBottomPercent ?? 100} onChange={(event) => updateBlock(index, { cropBottomPercent: percentOrDefault(event.target.value, 100) })} /></label><label>解答開始P<input type="number" value={block.answerPageStart ?? ''} onChange={(event) => updateBlock(index, { answerPageStart: numberOrUndefined(event.target.value) })} /></label><label>解説開始P<input type="number" value={block.explanationPageStart ?? ''} onChange={(event) => updateBlock(index, { explanationPageStart: numberOrUndefined(event.target.value) })} /></label></div><PdfQuestionPreview pdf={selectedPdf} page={block.problemPageStart ?? draft.problemPageStart} cropTopPercent={block.cropTopPercent} cropBottomPercent={block.cropBottomPercent} title="この設問で表示されるPDF範囲" compact /><div className="form-grid three-columns"><label>難易度<select value={blockDifficultyLabel(block.difficulty)} onChange={(event) => updateBlock(index, { difficulty: event.target.value as BlockDifficulty })}><option value="易しい">易しい</option><option value="標準">標準</option><option value="やや難">やや難</option><option value="難しい">難しい</option></select></label><label>補足<input value={block.description ?? ''} onChange={(event) => updateBlock(index, { description: event.target.value })} placeholder="例：商品、収益認識、現金預金" /></label><label>メモ<input value={block.memo ?? ''} onChange={(event) => updateBlock(index, { memo: event.target.value })} placeholder="教材本文は入れない" /></label></div></div>)}
          </div>
        </details>

        <div className="button-row sticky-actions"><button className="accent" onClick={() => { void saveMapping(); }}>保存する</button><button className="secondary" onClick={() => { void onOpenProblem(draft.problemId); }}>今すぐ解くで開く</button></div>
      </div>

      <div className="setup-preview-panel"><PdfViewerPanel pdf={selectedPdf} title={selectedPdf ? selectedPdf.title : 'PDF未選択'} label={previewKind === 'answer' ? '解答' : previewKind === 'explanation' ? '解説' : '問題'} page={page} pageStart={1} pageEnd={selectedPdf?.pageCount ?? 1} onPageChange={setPage} large /></div>

      <details className="panel mapping-list-panel" open><summary>保存済み設定 / 設問数</summary><div className="panel-body"><div className="progress-summary-grid"><div><strong>{mappings.length}</strong><span>保存済み紐付け</span></div><div><strong>{mapped.length}</strong><span>設定済み問題ID</span></div><div><strong>{unmapped.length}</strong><span>未設定問題ID</span></div></div><div className="table-wrap short-wrap"><table className="compact-table"><thead><tr><th>問題ID</th><th>論点</th><th>PDF</th><th>問題</th><th>解答</th><th>解説</th><th>設問</th><th>操作</th></tr></thead><tbody>{mappings.map((mapping) => { const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0]; const pdf = pdfs.find((item) => item.id === mapping.materialPdfId); return <tr key={mapping.id}><td>{problem.displayId}</td><td>{problem.topic}</td><td>{pdf?.title ?? 'PDF未登録'}</td><td>{mapping.problemPageStart}-{mapping.problemPageEnd}</td><td>{mapping.answerPageStart ? `${mapping.answerPageStart}-${mapping.answerPageEnd ?? mapping.answerPageStart}` : '-'}</td><td>{mapping.explanationPageStart ? `${mapping.explanationPageStart}-${mapping.explanationPageEnd ?? mapping.explanationPageStart}` : '-'}</td><td>{mapping.problemBlocks?.length ?? 1}</td><td><button onClick={() => { setDraft({ ...mapping, problemBlocks: ensureBlocks(mapping) }); setSelectedPdfId(mapping.materialPdfId); void onOpenProblem(mapping.problemId); }}>確認</button> <button className="danger" onClick={() => { void onDeleteMapping(mapping.id); }}>削除</button></td></tr>; })}{mappings.length === 0 && <tr><td colSpan={8}>まだページ紐付けがありません。</td></tr>}</tbody></table></div>{unmapped.length > 0 && <p className="notice-small">未設定問題ID例：{unmapped.slice(0, 16).map((problem) => `${problem.displayId} ${problem.topic}`).join(' / ')}{unmapped.length > 16 ? ' ...' : ''}</p>}</div></details>

      <details className="panel mapping-list-panel"><summary>PDF教材データの削除</summary><div className="panel-body"><div className="table-wrap short-wrap"><table className="compact-table"><thead><tr><th>教材名</th><th>ファイル</th><th>ページ数</th><th>サイズ</th><th>操作</th></tr></thead><tbody>{pdfs.map((pdf) => <tr key={pdf.id}><td>{pdf.title}</td><td>{pdf.fileName}</td><td>{pdf.pageCount}</td><td>{formatFileSize(pdf.size)}</td><td><button className="danger" onClick={() => { if (window.confirm('このPDF教材データと問題IDへの紐付けを削除します。答案履歴・白紙再現カードは削除されません。本当に削除しますか？')) void onDeletePdf(pdf.id); }}>削除</button></td></tr>)}{pdfs.length === 0 && <tr><td colSpan={5}>PDF教材は未登録です。</td></tr>}</tbody></table></div></div></details>
    </section>
  );
}
