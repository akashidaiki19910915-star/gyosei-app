import { useEffect, useMemo, useRef, useState } from 'react';
import { problemCatalog } from '../data/problemCatalog';
import type { MaterialPdf, MaterialPdfMapping, TimeMode } from '../types';
import { nowIso } from '../utils/dates';
import { timeModes } from '../utils/disposableStudy';
import { formatFileSize, getPdfPageCount, renderPdfPageToCanvas } from '../utils/pdfRenderer';

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
    createdAt: now,
    updatedAt: now,
  };
}

function numberOrUndefined(value: string): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

export function MaterialSetupWorkspace({ pdfs, mappings, selectedProblemId, onSavePdf, onDeletePdf, onSaveMapping, onDeleteMapping, onOpenProblem, onMessage }: Props) {
  const [title, setTitle] = useState('');
  const [sourceName, setSourceName] = useState('CPA問題集');
  const [bookType, setBookType] = useState<MaterialPdf['bookType']>('商業簿記');
  const [selectedPdfId, setSelectedPdfId] = useState(pdfs[0]?.id ?? '');
  const [draft, setDraft] = useState<MaterialPdfMapping>(() => blankMapping(selectedProblemId, pdfs[0]?.id ?? ''));
  const [previewKind, setPreviewKind] = useState<RangeKind>('problem');
  const [page, setPage] = useState(1);
  const [scale, setScale] = useState(1.05);
  const [loading, setLoading] = useState(false);
  const [renderError, setRenderError] = useState('');
  const canvasRef = useRef<HTMLCanvasElement | null>(null);

  const selectedPdf = useMemo(() => pdfs.find((pdf) => pdf.id === selectedPdfId) ?? pdfs[0], [pdfs, selectedPdfId]);
  const mappedProblemIds = useMemo(() => new Set(mappings.map((mapping) => mapping.problemId)), [mappings]);
  const unmapped = problemCatalog.filter((problem) => !mappedProblemIds.has(problem.id));
  const mapped = problemCatalog.filter((problem) => mappedProblemIds.has(problem.id));

  useEffect(() => {
    if (!selectedPdfId && pdfs[0]) setSelectedPdfId(pdfs[0].id);
  }, [pdfs, selectedPdfId]);

  useEffect(() => {
    const existing = mappings.find((mapping) => mapping.problemId === selectedProblemId) ?? blankMapping(selectedProblemId, selectedPdf?.id ?? '');
    setDraft({ ...existing, materialPdfId: existing.materialPdfId || selectedPdf?.id || '' });
    setSelectedPdfId(existing.materialPdfId || selectedPdf?.id || '');
  }, [selectedProblemId, mappings, selectedPdf?.id]);

  const activeRange = (() => {
    if (previewKind === 'answer' && draft.answerPageStart) return { start: draft.answerPageStart, end: draft.answerPageEnd ?? draft.answerPageStart };
    if (previewKind === 'explanation' && draft.explanationPageStart) return { start: draft.explanationPageStart, end: draft.explanationPageEnd ?? draft.explanationPageStart };
    return { start: draft.problemPageStart, end: draft.problemPageEnd };
  })();

  useEffect(() => {
    if (!selectedPdf || !canvasRef.current) return;
    let cancelled = false;
    setLoading(true);
    setRenderError('');
    const safePage = Math.max(1, Math.min(page, selectedPdf.pageCount || page));
    renderPdfPageToCanvas(selectedPdf.pdfBlob, safePage, scale, canvasRef.current)
      .catch(() => {
        if (!cancelled) setRenderError('PDFページの表示に失敗しました。保護PDF、破損ファイル、またはブラウザ非対応の可能性があります。');
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => { cancelled = true; };
  }, [selectedPdf, page, scale]);

  useEffect(() => {
    setPage(activeRange.start || 1);
  }, [previewKind, draft.problemPageStart, draft.answerPageStart, draft.explanationPageStart]);

  function update(patch: Partial<MaterialPdfMapping>) {
    const nextProblemId = patch.problemId ?? draft.problemId;
    const problem = problemCatalog.find((item) => item.id === nextProblemId) ?? problemCatalog[0];
    setDraft((current) => ({ ...current, ...patch, problemId: nextProblemId, displayId: problem.displayId, updatedAt: nowIso() }));
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
      onMessage('PDFを端末内IndexedDBへ保存しました。続けてページを見ながら問題IDへ紐づけてください。');
    } catch {
      onMessage('PDFの読み込みに失敗しました。破損ファイル、保護PDF、ブラウザ非対応の可能性があります。');
    }
  }

  async function saveMapping() {
    if (!draft.materialPdfId) {
      onMessage('先に対象PDFを選択してください');
      return;
    }
    await onSaveMapping({ ...draft, updatedAt: nowIso() });
  }

  function jumpToRange(kind: RangeKind) {
    setPreviewKind(kind);
    if (kind === 'answer' && draft.answerPageStart) setPage(draft.answerPageStart);
    else if (kind === 'explanation' && draft.explanationPageStart) setPage(draft.explanationPageStart);
    else setPage(draft.problemPageStart);
  }

  return (
    <section className="material-setup-grid">
      <div className="panel setup-control-panel mode-card">
        <p className="eyebrow">教材を設定する</p>
        <h2>PDFを見ながらページ範囲を登録</h2>
        <p className="notice-small">PDF本体・教材本文・解答解説はGitHubや外部サーバーへ送信しません。利用者本人の端末内IndexedDBに保存します。</p>

        <div className="form-grid three-columns">
          <label>表示名<input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="例：商業簿記 試験対策編" /></label>
          <label>教材区分
            <select value={bookType} onChange={(event) => setBookType(event.target.value as MaterialPdf['bookType'])}>
              <option value="商業簿記">商業簿記</option>
              <option value="工業簿記">工業簿記</option>
              <option value="その他">その他</option>
            </select>
          </label>
          <label>出所メモ<input value={sourceName} onChange={(event) => setSourceName(event.target.value)} placeholder="例：正規取得PDF" /></label>
        </div>
        <label className="import-label pdf-import-label">PDFを選択してプレビュー
          <input type="file" accept="application/pdf" onChange={(event) => { void importPdf(event.target.files?.[0]); }} />
        </label>

        <div className="form-grid two-columns">
          <label>対象PDF
            <select value={draft.materialPdfId || selectedPdfId} onChange={(event) => { setSelectedPdfId(event.target.value); update({ materialPdfId: event.target.value }); }}>
              <option value="">PDFを選択</option>
              {pdfs.map((pdf) => <option key={pdf.id} value={pdf.id}>{pdf.title}（{pdf.pageCount}P）</option>)}
            </select>
          </label>
          <label>問題ID
            <select value={draft.problemId} onChange={(event) => { const existing = mappings.find((mapping) => mapping.problemId === event.target.value); setDraft(existing ?? blankMapping(event.target.value, draft.materialPdfId || selectedPdfId)); }}>
              {problemCatalog.map((problem) => <option key={problem.id} value={problem.id}>{problem.sectionLabel} {problem.displayId} {problem.topic}</option>)}
            </select>
          </label>
        </div>

        <div className="form-grid six-columns">
          <label>問題開始P<input type="number" min="1" value={draft.problemPageStart} onChange={(event) => update({ problemPageStart: Number(event.target.value) || 1 })} /></label>
          <label>問題終了P<input type="number" min="1" value={draft.problemPageEnd} onChange={(event) => update({ problemPageEnd: Number(event.target.value) || draft.problemPageStart })} /></label>
          <label>解答開始P<input type="number" min="1" value={draft.answerPageStart ?? ''} onChange={(event) => update({ answerPageStart: numberOrUndefined(event.target.value) })} /></label>
          <label>解答終了P<input type="number" min="1" value={draft.answerPageEnd ?? ''} onChange={(event) => update({ answerPageEnd: numberOrUndefined(event.target.value) })} /></label>
          <label>解説開始P<input type="number" min="1" value={draft.explanationPageStart ?? ''} onChange={(event) => update({ explanationPageStart: numberOrUndefined(event.target.value) })} /></label>
          <label>解説終了P<input type="number" min="1" value={draft.explanationPageEnd ?? ''} onChange={(event) => update({ explanationPageEnd: numberOrUndefined(event.target.value) })} /></label>
        </div>

        <div className="form-grid three-columns">
          <label>想定時間
            <select value={draft.estimatedMinutes} onChange={(event) => update({ estimatedMinutes: event.target.value as TimeMode })}>
              {timeModes.map((mode) => <option key={mode} value={mode}>{mode}</option>)}
            </select>
          </label>
          <label>難易度
            <select value={draft.difficulty} onChange={(event) => update({ difficulty: event.target.value as MaterialPdfMapping['difficulty'] })}>
              <option value="易">易</option><option value="標準">標準</option><option value="難">難</option>
            </select>
          </label>
          <label>メモ<input value={draft.memo} onChange={(event) => update({ memo: event.target.value })} placeholder="教材本文は入れない" /></label>
        </div>

        <div className="button-row sticky-actions">
          <button className="accent" onClick={() => { void saveMapping(); }}>この問題IDに紐づけて保存</button>
          <button className="secondary" onClick={() => { void onOpenProblem(draft.problemId); }}>この問題を開く</button>
        </div>
      </div>

      <div className="panel setup-preview-panel mode-card">
        <div className="section-heading">
          <div>
            <p className="eyebrow">プレビュー：{previewKind === 'answer' ? '解答' : previewKind === 'explanation' ? '解説' : '問題'}</p>
            <h2>{selectedPdf ? selectedPdf.title : 'PDF未選択'}</h2>
            {selectedPdf && <p>{selectedPdf.fileName} / {selectedPdf.pageCount}P / {formatFileSize(selectedPdf.size)}</p>}
          </div>
          <div className="button-row">
            <button className={previewKind === 'problem' ? 'accent' : 'secondary'} onClick={() => jumpToRange('problem')}>問題</button>
            <button className={previewKind === 'answer' ? 'accent' : 'secondary'} disabled={!draft.answerPageStart} onClick={() => jumpToRange('answer')}>解答</button>
            <button className={previewKind === 'explanation' ? 'accent' : 'secondary'} disabled={!draft.explanationPageStart} onClick={() => jumpToRange('explanation')}>解説</button>
          </div>
        </div>
        <div className="pdf-toolbar">
          <button disabled={!selectedPdf} onClick={() => setPage((current) => Math.max(1, current - 1))}>前ページ</button>
          <label>ページ<input type="number" min="1" max={selectedPdf?.pageCount ?? 1} value={page} onChange={(event) => setPage(Math.max(1, Number(event.target.value) || 1))} /></label>
          <span>登録範囲：{activeRange.start}〜{activeRange.end}ページ</span>
          <button disabled={!selectedPdf} onClick={() => setPage((current) => Math.min(selectedPdf?.pageCount ?? current + 1, current + 1))}>次ページ</button>
          <button onClick={() => setScale((current) => Math.max(0.7, current - 0.15))}>縮小</button>
          <button onClick={() => setScale((current) => Math.min(2.2, current + 0.15))}>拡大</button>
        </div>
        {renderError && <p className="warning-text">{renderError}</p>}
        <div className="pdf-canvas-wrap setup-canvas-wrap">
          {!selectedPdf && <p className="empty">左側からPDFを選択してください。</p>}
          {loading && <p className="empty">PDFを表示中...</p>}
          <canvas ref={canvasRef} className="pdf-canvas" />
        </div>
      </div>

      <details className="panel mapping-list-panel" open>
        <summary>保存済みマッピング / 設定状況</summary>
        <div className="panel-body">
          <div className="progress-summary-grid">
            <div><strong>{mappings.length}</strong><span>保存済み紐付け</span></div>
            <div><strong>{mapped.length}</strong><span>設定済み問題ID</span></div>
            <div><strong>{unmapped.length}</strong><span>未設定問題ID</span></div>
          </div>
          <div className="table-wrap short-wrap">
            <table className="compact-table">
              <thead><tr><th>問題ID</th><th>論点</th><th>PDF</th><th>問題</th><th>解答</th><th>解説</th><th>操作</th></tr></thead>
              <tbody>
                {mappings.map((mapping) => {
                  const problem = problemCatalog.find((item) => item.id === mapping.problemId) ?? problemCatalog[0];
                  const pdf = pdfs.find((item) => item.id === mapping.materialPdfId);
                  return (
                    <tr key={mapping.id}>
                      <td>{problem.displayId}</td><td>{problem.topic}</td><td>{pdf?.title ?? 'PDF未登録'}</td>
                      <td>{mapping.problemPageStart}-{mapping.problemPageEnd}</td>
                      <td>{mapping.answerPageStart ? `${mapping.answerPageStart}-${mapping.answerPageEnd ?? mapping.answerPageStart}` : '-'}</td>
                      <td>{mapping.explanationPageStart ? `${mapping.explanationPageStart}-${mapping.explanationPageEnd ?? mapping.explanationPageStart}` : '-'}</td>
                      <td><button onClick={() => { setDraft(mapping); setSelectedPdfId(mapping.materialPdfId); void onOpenProblem(mapping.problemId); }}>確認</button> <button className="danger" onClick={() => { void onDeleteMapping(mapping.id); }}>削除</button></td>
                    </tr>
                  );
                })}
                {mappings.length === 0 && <tr><td colSpan={7}>まだページ紐付けがありません。</td></tr>}
              </tbody>
            </table>
          </div>
          {unmapped.length > 0 && <p className="notice-small">未設定問題ID例：{unmapped.slice(0, 16).map((problem) => `${problem.displayId} ${problem.topic}`).join(' / ')}{unmapped.length > 16 ? ' ...' : ''}</p>}
        </div>
      </details>

      <details className="panel mapping-list-panel">
        <summary>PDF教材データの削除</summary>
        <div className="panel-body">
          <div className="table-wrap short-wrap">
            <table className="compact-table">
              <thead><tr><th>教材名</th><th>ファイル</th><th>ページ数</th><th>サイズ</th><th>操作</th></tr></thead>
              <tbody>
                {pdfs.map((pdf) => <tr key={pdf.id}><td>{pdf.title}</td><td>{pdf.fileName}</td><td>{pdf.pageCount}</td><td>{formatFileSize(pdf.size)}</td><td><button className="danger" onClick={() => { if (window.confirm('このPDF教材データと問題IDへの紐付けを削除します。答案履歴・白紙再現カードは削除されません。本当に削除しますか？')) void onDeletePdf(pdf.id); }}>削除</button></td></tr>)}
                {pdfs.length === 0 && <tr><td colSpan={5}>PDF教材は未登録です。</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </details>
    </section>
  );
}
