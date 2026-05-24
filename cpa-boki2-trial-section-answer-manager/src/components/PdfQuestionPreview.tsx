import { useEffect, useRef, useState } from 'react';
import type { MaterialPdf } from '../types';
import { renderPdfPageToCanvas } from '../utils/pdfRenderer';

interface Props {
  pdf?: MaterialPdf;
  page?: number;
  cropTopPercent?: number;
  cropBottomPercent?: number;
  title?: string;
  compact?: boolean;
}

function clampPercent(value: number | undefined, fallback: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(0, Math.min(100, value));
}

export function PdfQuestionPreview({ pdf, page, cropTopPercent, cropBottomPercent, title = 'PDFの該当部分', compact = false }: Props) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [canvasHeight, setCanvasHeight] = useState(0);
  const [canvasWidth, setCanvasWidth] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const top = clampPercent(cropTopPercent, 0);
  const rawBottom = clampPercent(cropBottomPercent, 100);
  const bottom = Math.max(top + 8, rawBottom);
  const safePage = page && page > 0 ? page : 1;
  const cropHeight = canvasHeight > 0 ? Math.max(120, Math.round(canvasHeight * ((bottom - top) / 100))) : compact ? 180 : 240;
  const translateY = canvasHeight > 0 ? -Math.round(canvasHeight * (top / 100)) : 0;

  useEffect(() => {
    if (!pdf || !canvasRef.current) return;
    let cancelled = false;
    setLoading(true);
    setError('');
    renderPdfPageToCanvas(pdf.pdfBlob, Math.min(safePage, pdf.pageCount), compact ? 0.9 : 1.05, canvasRef.current)
      .then(() => {
        if (!cancelled && canvasRef.current) {
          setCanvasHeight(canvasRef.current.height);
          setCanvasWidth(canvasRef.current.width);
        }
      })
      .catch(() => {
        if (!cancelled) setError('PDFプレビューを表示できませんでした。');
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => { cancelled = true; };
  }, [pdf, safePage, compact]);

  if (!pdf) {
    return <div className="question-pdf-preview empty-preview"><strong>{title}</strong><p>この設問にPDFが紐づいていません。教材設定でPDFページと表示範囲を登録してください。</p></div>;
  }

  return (
    <div className={`question-pdf-preview ${compact ? 'compact-preview' : ''}`}>
      <div className="question-pdf-preview-header">
        <strong>{title}</strong>
        <span>P{Math.min(safePage, pdf.pageCount)} / {pdf.pageCount}　表示範囲：上{top}%〜{bottom}%</span>
      </div>
      {loading && <p className="empty">PDFを表示中...</p>}
      {error && <p className="warning-text">{error}</p>}
      <div className="question-pdf-crop-window" style={{ height: cropHeight }}>
        <canvas
          ref={canvasRef}
          className="question-pdf-canvas"
          style={{ transform: `translateY(${translateY}px)`, width: canvasWidth ? '100%' : undefined }}
        />
      </div>
    </div>
  );
}
