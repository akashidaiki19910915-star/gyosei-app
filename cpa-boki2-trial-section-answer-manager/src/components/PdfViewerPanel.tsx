import { useEffect, useRef, useState } from 'react';
import type { MaterialPdf } from '../types';
import { renderPdfPageToCanvas } from '../utils/pdfRenderer';

interface Props {
  pdf?: MaterialPdf;
  title: string;
  page: number;
  pageStart?: number;
  pageEnd?: number;
  onPageChange: (page: number) => void;
  onClose?: () => void;
  label?: string;
  large?: boolean;
}

type FitMode = 'custom' | 'width' | 'height';

function clamp(value: number, start: number, end: number): number {
  return Math.max(start, Math.min(end, value));
}

export function PdfViewerPanel({ pdf, title, page, pageStart, pageEnd, onPageChange, onClose, label = '問題', large = false }: Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const [scale, setScale] = useState(large ? 1.35 : 1.1);
  const [fitMode, setFitMode] = useState<FitMode>('custom');
  const [full, setFull] = useState(false);
  const [loading, setLoading] = useState(false);
  const [renderError, setRenderError] = useState('');

  const start = pageStart ?? 1;
  const end = pageEnd ?? pdf?.pageCount ?? start;
  const safePage = clamp(page, start, end);

  useEffect(() => {
    if (!pdf || !canvasRef.current) return;
    let cancelled = false;
    setLoading(true);
    setRenderError('');
    renderPdfPageToCanvas(pdf.pdfBlob, safePage, scale, canvasRef.current)
      .catch(() => {
        if (!cancelled) setRenderError('PDFページの表示に失敗しました。保護PDF、破損ファイル、またはブラウザ非対応の可能性があります。');
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => { cancelled = true; };
  }, [pdf, safePage, scale]);

  useEffect(() => {
    if (!canvasRef.current || !wrapRef.current) return;
    if (fitMode === 'custom') return;
    const canvas = canvasRef.current;
    const wrap = wrapRef.current;
    if (!canvas.width || !canvas.height) return;
    const horizontalPadding = 32;
    const verticalPadding = 32;
    if (fitMode === 'width') setScale((current) => Math.max(0.55, Math.min(2.8, current * ((wrap.clientWidth - horizontalPadding) / canvas.clientWidth))));
    if (fitMode === 'height') setScale((current) => Math.max(0.55, Math.min(2.8, current * ((wrap.clientHeight - verticalPadding) / canvas.clientHeight))));
  }, [fitMode, loading]);

  const movePage = (next: number) => onPageChange(clamp(next, start, end));

  return (
    <section className={`panel pdf-viewer-panel ${large ? 'pdf-viewer-large' : ''} ${full ? 'pdf-viewer-full' : ''}`}>
      <div className="section-heading pdf-viewer-heading">
        <div>
          <p className="eyebrow">{label}PDF</p>
          <h2>{title}</h2>
          {pdf && <p>現在ページ {safePage} / 全{pdf.pageCount}ページ　登録範囲：{start}〜{end}ページ　倍率：{Math.round(scale * 100)}%</p>}
        </div>
        <div className="button-row">
          <button onClick={() => setFull((current) => !current)}>{full ? '通常表示' : 'フルスクリーン風表示'}</button>
          {onClose && <button className="secondary" onClick={onClose}>閉じる</button>}
        </div>
      </div>
      <div className="pdf-toolbar pdf-toolbar-strong">
        <button disabled={!pdf} onClick={() => movePage(safePage - 1)}>前ページ</button>
        <label>ページ番号<input type="number" min={start} max={end} value={safePage} onChange={(event) => movePage(Number(event.target.value) || start)} /></label>
        <button disabled={!pdf} onClick={() => movePage(safePage + 1)}>次ページ</button>
        <button onClick={() => { setFitMode('custom'); setScale((current) => Math.max(0.55, current - 0.15)); }}>縮小</button>
        <button onClick={() => { setFitMode('custom'); setScale((current) => Math.min(2.8, current + 0.15)); }}>拡大</button>
        <button onClick={() => { setFitMode('custom'); setScale(1); }}>100%</button>
        <button onClick={() => setFitMode('width')}>幅に合わせる</button>
        <button onClick={() => setFitMode('height')}>高さに合わせる</button>
      </div>
      {renderError && <p className="warning-text">{renderError}</p>}
      <div ref={wrapRef} className="pdf-canvas-wrap pdf-reader-wrap">
        {!pdf && <p className="empty">PDFが選択されていません。</p>}
        {loading && <p className="empty">PDFを表示中...</p>}
        <canvas ref={canvasRef} className="pdf-canvas" />
      </div>
    </section>
  );
}
